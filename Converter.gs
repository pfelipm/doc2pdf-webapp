/**
 * Converter.gs
 * Motor de conversión nativo basado en Google Drive API v3 y UrlFetchApp.
 * Soporta subida directa, subida fragmentada (Resumable Upload), exportación PDF
 * sin límite de 10 MB, descarga fragmentada y conversión desde enlace de Drive
 * con defensa anti-Confused Deputy y verificación estricta de propiedad de sesión.
 */

/**
 * Genera un identificador único de trabajo legible para auditoría y trazabilidad.
 */
function generateJobId_() {
  let stamp = '';
  if (typeof Utilities !== 'undefined' && Utilities.formatDate) {
    stamp = Utilities.formatDate(new Date(), 'Europe/Madrid', 'yyyyMMdd-HHmmss');
  } else {
    stamp = String(Date.now());
  }
  const rand = Math.floor(Math.random() * 0xffff)
    .toString(16)
    .toUpperCase()
    .padStart(4, '0');
  return 'JOB-' + stamp + '-' + rand;
}

/**
 * Extrae la extensión en minúsculas sin punto de un nombre de archivo.
 */
function extractExtension_(fileName) {
  const name = String(fileName || '').trim();
  const idx = name.lastIndexOf('.');
  if (idx <= 0 || idx === name.length - 1) return '';
  return name.slice(idx + 1).toLowerCase();
}

/**
 * Construye el nombre del archivo PDF resultante manteniendo el nombre base.
 */
function buildPdfFileName_(originalName) {
  const raw = String(originalName || 'documento').trim();
  const idx = raw.lastIndexOf('.');
  if (idx > 0 && idx < raw.length - 1 && raw.length - idx <= 7) {
    return raw.slice(0, idx) + '.pdf';
  }
  return raw + '.pdf';
}

/**
 * Extrae el ID de archivo desde distintos formatos de URL de Google Drive/Docs o ID directo.
 * @return {string}
 */
function extractDriveFileId_(input) {
  const raw = String(input || '').trim();
  if (!raw) return '';

  const pathMatch = raw.match(/\/d\/([a-zA-Z0-9_-]{15,})/);
  if (pathMatch && pathMatch[1]) {
    return pathMatch[1];
  }

  const queryMatch = raw.match(/[?&]id=([a-zA-Z0-9_-]{15,})/);
  if (queryMatch && queryMatch[1]) {
    return queryMatch[1];
  }

  const directMatch = raw.match(/^([a-zA-Z0-9_-]{15,})$/);
  if (directMatch && directMatch[1]) {
    return directMatch[1];
  }

  return '';
}

/**
 * Valida y normaliza las opciones de exportación PDF para hojas de cálculo.
 */
function normalizeSheetOptions_(opts) {
  const input = opts || {};
  const size = input.size === 'letter' ? 'letter' : 'A4';
  const portrait = input.portrait === true;
  const rawScale = String(input.scale || '2');
  const scale = ['1', '2', '4'].indexOf(rawScale) !== -1 ? rawScale : '2';
  const gridlines = input.gridlines === false ? false : true;
  const margins = input.margins === 'narrow' ? 'narrow' : 'normal';

  return {
    size: size,
    portrait: portrait,
    scale: scale,
    gridlines: gridlines,
    margins: margins
  };
}

/**
 * Resume las opciones de hoja de cálculo para registrarlas en Auditoría.
 */
function formatSheetOptionsForAudit_(sheetOpts) {
  if (!sheetOpts) return '';
  const scaleLabels = {
    '1': 'Escala 100%',
    '2': 'Ajustar al ancho',
    '4': 'Ajustar a 1 página'
  };
  return [
    sheetOpts.size === 'letter' ? 'Carta' : 'A4',
    sheetOpts.portrait ? 'Vertical' : 'Horizontal',
    scaleLabels[sheetOpts.scale] || 'Ajustar al ancho',
    sheetOpts.gridlines ? 'Cuadrícula: Sí' : 'Cuadrícula: No',
    sheetOpts.margins === 'narrow' ? 'Márgenes: Estrechos' : 'Márgenes: Normales'
  ].join(' | ');
}

/**
 * Construye la URL oficial de exportación PDF de Google Workspace para evitar el límite de 10 MB
 * de Drive.Files.export y permitir opciones de maquetación en hojas de cálculo.
 */
function buildExportUrl_(googleFileId, exportType, sheetOptions) {
  const safeId = encodeURIComponent(String(googleFileId));
  if (exportType === 'spreadsheets') {
    const opts = normalizeSheetOptions_(sheetOptions);
    const marginVal = opts.margins === 'narrow' ? '0.25' : '0.50';
    const params = [
      'format=pdf',
      'size=' + encodeURIComponent(opts.size),
      'portrait=' + (opts.portrait ? 'true' : 'false'),
      'scale=' + encodeURIComponent(opts.scale),
      'gridlines=' + (opts.gridlines ? 'true' : 'false'),
      'top_margin=' + marginVal,
      'bottom_margin=' + marginVal,
      'left_margin=' + marginVal,
      'right_margin=' + marginVal,
      'printtitle=false',
      'sheetnames=false',
      'fzr=true'
    ];
    return 'https://docs.google.com/spreadsheets/d/' + safeId + '/export?' + params.join('&');
  }

  if (exportType === 'presentation') {
    return 'https://docs.google.com/presentation/d/' + safeId + '/export?format=pdf';
  }

  return 'https://docs.google.com/document/d/' + safeId + '/export?format=pdf';
}

/**
 * Obtiene una cabecera HTTP de respuesta sin distinguir mayúsculas/minúsculas.
 */
function getCaseInsensitiveHeader_(headers, headerName) {
  if (!headers) return '';
  const target = String(headerName).toLowerCase();
  const keys = Object.keys(headers);
  for (let i = 0; i < keys.length; i++) {
    if (String(keys[i]).toLowerCase() === target) {
      const val = headers[keys[i]];
      return Array.isArray(val) ? String(val[0] || '') : String(val || '');
    }
  }
  return '';
}

/**
 * Exige que el solicitante actual sea el propietario legítimo de la sesión de subida/descarga.
 */
function assertSessionOwnership_(sessionObj, actionName) {
  const currentEmail = getRealActiveEmail_();
  if (!currentEmail || !sessionObj || !sessionObj.userEmail || sessionObj.userEmail !== currentEmail) {
    const err = new Error('No autorizado para operar sobre esta sesión de ' + (actionName || 'conversión') + '.');
    err.code = 'SESSION_FORBIDDEN';
    throw err;
  }
}

/**
 * Mitigación anti-Confused Deputy: impide usar el token del administrador para leer
 * la hoja de control, temporales internos o archivos privados no compartidos con el usuario.
 */
function assertDriveFileSharedWithRequester_(meta, userCtx, tempFolderId) {
  const ss = getSpreadsheet_();
  const ssId = ss && typeof ss.getId === 'function' ? ss.getId() : '';

  if (ssId && meta.id === ssId) {
    const err = new Error('No está permitido acceder a la hoja de cálculo de control del servicio.');
    err.code = 'DRIVE_ACCESS_FORBIDDEN';
    throw err;
  }

  if (tempFolderId && Array.isArray(meta.parents) && meta.parents.indexOf(tempFolderId) !== -1) {
    const err = new Error('No está permitido acceder a archivos temporales internos del servicio.');
    err.code = 'DRIVE_ACCESS_FORBIDDEN';
    throw err;
  }

  const requesterEmail = (userCtx.email || '').toLowerCase();
  const requesterDomain = (userCtx.domain || '').toLowerCase();

  const owners = Array.isArray(meta.owners) ? meta.owners : [];
  for (let i = 0; i < owners.length; i++) {
    const ownerEmail = owners[i] && owners[i].emailAddress ? String(owners[i].emailAddress).trim().toLowerCase() : '';
    if (ownerEmail && ownerEmail === requesterEmail) {
      return;
    }
  }

  const permissions = Array.isArray(meta.permissions) ? meta.permissions : [];
  for (let i = 0; i < permissions.length; i++) {
    const p = permissions[i];
    if (!p || !p.type) continue;
    const pType = String(p.type).toLowerCase();
    if (pType === 'anyone') {
      return;
    }
    if (pType === 'domain') {
      const pDomain = p.domain ? String(p.domain).trim().toLowerCase() : '';
      if (pDomain && requesterDomain && pDomain === requesterDomain) {
        return;
      }
    }
    if (pType === 'user') {
      const pEmail = p.emailAddress ? String(p.emailAddress).trim().toLowerCase() : '';
      if (pEmail && pEmail === requesterEmail) {
        return;
      }
    }
  }

  const err = new Error('Por seguridad, solo se pueden convertir archivos de Drive compartidos contigo o con tu dominio Workspace (no archivos privados de la cuenta del servicio).');
  err.code = 'DRIVE_ACCESS_NOT_SHARED';
  throw err;
}

/**
 * Exporta un documento nativo de Google Workspace a Blob PDF mediante UrlFetchApp.
 */
function exportGoogleFileToPdfBlob_(googleFileId, exportType, pdfFileName, sheetOptions) {
  const exportUrl = buildExportUrl_(googleFileId, exportType, sheetOptions);
  const retryableStatuses = [429, 500, 502, 503];
  let response = null;
  let code = 0;

  for (let attempt = 0; attempt < 2; attempt++) {
    response = UrlFetchApp.fetch(exportUrl, {
      method: 'get',
      headers: {
        Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
      },
      muteHttpExceptions: true
    });
    code = response.getResponseCode();
    if (code === 200 || retryableStatuses.indexOf(code) === -1 || attempt === 1) {
      break;
    }
    Utilities.sleep(900);
  }

  if (code !== 200) {
    throw new Error('Google Workspace no pudo exportar el archivo a PDF (HTTP ' + code + '). Comprueba que el documento no esté protegido ni dañado.');
  }

  const headers = response.getAllHeaders ? response.getAllHeaders() : response.getHeaders();
  const contentType = getCaseInsensitiveHeader_(headers, 'content-type').toLowerCase();
  if (contentType && contentType.indexOf('application/pdf') === -1) {
    throw new Error('La respuesta de exportación no devolvió un PDF válido (' + contentType + ').');
  }

  return response.getBlob().setName(pdfFileName);
}

/**
 * Prepara la entrega del PDF generado al navegador: directa en Base64 si <= 7 MB,
 * o mediante archivo temporal en Drive y descarga por bloques si > 7 MB.
 */
function preparePdfTransfer_(pdfBlob, pdfFileName, jobId, userEmail, tempFolderId) {
  const bytes = pdfBlob.getBytes();
  const pdfSizeBytes = bytes.length;

  if (pdfSizeBytes <= DIRECT_DOWNLOAD_MAX_BYTES) {
    return {
      transferMode: 'direct',
      pdfBase64: Utilities.base64Encode(bytes),
      pdfFileName: pdfFileName,
      pdfSizeBytes: pdfSizeBytes
    };
  }

  const outFile = Drive.Files.create(
    {
      name: '__out_' + jobId + '.pdf',
      mimeType: 'application/pdf',
      parents: [tempFolderId]
    },
    pdfBlob,
    { fields: 'id', supportsAllDrives: true }
  );

  const token = Utilities.getUuid();
  const totalChunks = Math.ceil(pdfSizeBytes / DOWNLOAD_CHUNK_BYTES);
  const dlSession = {
    fileId: outFile.id,
    totalBytes: pdfSizeBytes,
    totalChunks: totalChunks,
    pdfFileName: pdfFileName,
    userEmail: userEmail || ''
  };
  CacheService.getScriptCache().put('DOC2PDF_DL_' + token, JSON.stringify(dlSession), 600);

  return {
    transferMode: 'chunked',
    downloadToken: token,
    totalChunks: totalChunks,
    chunkSizeBytes: DOWNLOAD_CHUNK_BYTES,
    pdfFileName: pdfFileName,
    pdfSizeBytes: pdfSizeBytes
  };
}

/**
 * Convierte un archivo local pequeño/mediano (<= 6 MB) enviado en Base64 en una sola llamada.
 */
function convertLocalFileDirect_(payload) {
  const startedAt = Date.now();
  const config = getConfig_();
  const jobId = generateJobId_();
  const rawName = String((payload && payload.fileName) || 'documento').trim();
  const ext = extractExtension_(rawName);
  const format = getFormatByExtension_(ext, config);
  const userCtx = assertAuthorizedForJob_(config, 'Local', rawName, ext, format ? format.family : '');

  let tempGoogleId = '';
  let inputBytesCount = 0;
  const sheetOpts = format && format.exportType === 'spreadsheets'
    ? normalizeSheetOptions_(payload && payload.sheetOptions)
    : null;
  const sheetOptsAudit = sheetOpts ? formatSheetOptionsForAudit_(sheetOpts) : '';

  try {
    if (!payload || !payload.base64Data) {
      throw new Error('No se recibieron datos del archivo.');
    }
    if (!format) {
      throw new Error('Formato no soportado o familia deshabilitada en Ajustes: .' + (ext || '(sin extensión)'));
    }

    const bytes = Utilities.base64Decode(payload.base64Data);
    inputBytesCount = bytes.length;
    if (inputBytesCount === 0) {
      throw new Error('El archivo está vacío (0 bytes).');
    }
    const maxBytes = config.TAMANO_MAX_MB * 1024 * 1024;
    if (inputBytesCount > maxBytes) {
      throw new Error('El archivo supera el máximo permitido de ' + config.TAMANO_MAX_MB + ' MB.');
    }

    const tempFolderId = getOrCreateTempFolder_(config);
    const inputBlob = Utilities.newBlob(bytes, format.sourceMime, rawName);
    const created = Drive.Files.create(
      {
        name: '__tmp_' + jobId,
        mimeType: format.targetGoogleMime,
        parents: [tempFolderId]
      },
      inputBlob,
      { fields: 'id', supportsAllDrives: true }
    );
    tempGoogleId = created.id;

    const pdfFileName = buildPdfFileName_(rawName);
    const pdfBlob = exportGoogleFileToPdfBlob_(tempGoogleId, format.exportType, pdfFileName, sheetOpts);
    const transfer = preparePdfTransfer_(pdfBlob, pdfFileName, jobId, userCtx.email, tempFolderId);
    const durationMs = Date.now() - startedAt;

    incrementUserDailyUsage_(userCtx.email);
    logAuditJob_({
      jobId: jobId,
      userEmail: userCtx.displayEmail,
      origin: 'Local',
      fileName: rawName,
      ext: ext,
      family: format.family,
      inputBytes: inputBytesCount,
      pdfBytes: transfer.pdfSizeBytes,
      durationMs: durationMs,
      status: 'OK',
      detail: 'Conversión completada',
      sheetOptions: sheetOptsAudit
    }, config);

    return Object.assign({
      jobId: jobId,
      ext: ext,
      family: format.family,
      inputSizeBytes: inputBytesCount,
      durationMs: durationMs
    }, transfer);
  } catch (err) {
    const durationMs = Date.now() - startedAt;
    logAuditJob_({
      jobId: jobId,
      userEmail: userCtx.displayEmail,
      origin: 'Local',
      fileName: rawName,
      ext: ext,
      family: format ? format.family : '',
      inputBytes: inputBytesCount,
      pdfBytes: 0,
      durationMs: durationMs,
      status: 'ERROR',
      detail: err.message,
      sheetOptions: sheetOptsAudit
    }, config);
    throw err;
  } finally {
    if (tempGoogleId) {
      removeDriveFilePermanently_(tempGoogleId);
    }
  }
}

/**
 * Inicia una sesión Resumable Upload en Drive API v3 para archivos locales grandes.
 */
function initChunkedUpload_(payload) {
  const config = getConfig_();
  const jobId = generateJobId_();
  const rawName = String((payload && payload.fileName) || 'documento').trim();
  const totalBytes = Number(payload && payload.totalBytes) || 0;
  const ext = extractExtension_(rawName);
  const format = getFormatByExtension_(ext, config);
  const userCtx = assertAuthorizedForJob_(config, 'Local', rawName, ext, format ? format.family : '');

  if (!format) {
    throw new Error('Formato no soportado o familia deshabilitada: .' + (ext || '(sin extensión)'));
  }
  const maxBytes = config.TAMANO_MAX_MB * 1024 * 1024;
  if (totalBytes <= 0 || totalBytes > maxBytes) {
    throw new Error('Tamaño inválido o superior al máximo de ' + config.TAMANO_MAX_MB + ' MB.');
  }

  const tempFolderId = getOrCreateTempFolder_(config);
  const initResp = UrlFetchApp.fetch(
    'https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable&fields=id&supportsAllDrives=true',
    {
      method: 'post',
      contentType: 'application/json; charset=UTF-8',
      headers: {
        Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
        'X-Upload-Content-Type': format.sourceMime,
        'X-Upload-Content-Length': String(totalBytes)
      },
      payload: JSON.stringify({
        name: '__tmp_' + jobId,
        mimeType: format.targetGoogleMime,
        parents: [tempFolderId]
      }),
      muteHttpExceptions: true
    }
  );

  const initCode = initResp.getResponseCode();
  if (initCode < 200 || initCode >= 300) {
    throw new Error('No se pudo iniciar la subida fragmentada a Drive (HTTP ' + initCode + ').');
  }

  const headers = initResp.getAllHeaders ? initResp.getAllHeaders() : initResp.getHeaders();
  const sessionUrl = getCaseInsensitiveHeader_(headers, 'location');
  if (!sessionUrl) {
    throw new Error('Drive API no devolvió URL de sesión para la subida fragmentada.');
  }

  const uploadId = Utilities.getUuid();
  const sheetOpts = format.exportType === 'spreadsheets'
    ? normalizeSheetOptions_(payload && payload.sheetOptions)
    : null;

  const sessionObj = {
    jobId: jobId,
    uploadId: uploadId,
    sessionUrl: sessionUrl,
    fileName: rawName,
    ext: ext,
    family: format.family,
    exportType: format.exportType,
    sourceMime: format.sourceMime,
    totalBytes: totalBytes,
    tempFolderId: tempFolderId,
    userEmail: userCtx.email,
    displayEmail: userCtx.displayEmail,
    sheetOptions: sheetOpts,
    startedAtMs: Date.now(),
    googleFileId: ''
  };

  CacheService.getScriptCache().put('DOC2PDF_UP_' + uploadId, JSON.stringify(sessionObj), 900);
  return {
    uploadId: uploadId,
    jobId: jobId,
    chunkSizeBytes: CHUNK_SIZE_BYTES
  };
}

/**
 * Envía un fragmento de bytes a la sesión Resumable Upload de Drive API v3.
 */
function uploadFileChunk_(payload) {
  const uploadId = String((payload && payload.uploadId) || '');
  const cache = CacheService.getScriptCache();
  const rawSession = cache.get('DOC2PDF_UP_' + uploadId);
  if (!rawSession) {
    throw new Error('La sesión de subida ha expirado. Vuelve a intentarlo.');
  }

  const session = JSON.parse(rawSession);
  assertSessionOwnership_(session, 'subida');

  const chunkBytes = Utilities.base64Decode(payload.base64Chunk || '');
  const start = Number(payload.start);
  const end = start + chunkBytes.length - 1;
  if (chunkBytes.length === 0 || isNaN(start) || start < 0 || end >= session.totalBytes) {
    throw new Error('Rango de fragmento inválido.');
  }

  const putResp = UrlFetchApp.fetch(session.sessionUrl, {
    method: 'put',
    contentType: session.sourceMime,
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
      'Content-Range': 'bytes ' + start + '-' + end + '/' + session.totalBytes
    },
    payload: chunkBytes,
    muteHttpExceptions: true
  });

  const code = putResp.getResponseCode();
  if (code === 308) {
    return { completed: false, nextOffset: end + 1 };
  }

  if (code === 200 || code === 201) {
    const body = JSON.parse(putResp.getContentText() || '{}');
    if (!body.id) {
      throw new Error('Drive completó la subida pero no devolvió ID de documento.');
    }
    session.googleFileId = body.id;
    cache.put('DOC2PDF_UP_' + uploadId, JSON.stringify(session), 900);
    return { completed: true, googleFileId: body.id };
  }

  throw new Error('Error enviando fragmento a Drive (HTTP ' + code + ').');
}

/**
 * Finaliza una conversión fragmentada: valida propiedad de sesión, exporta a PDF, audita y purga el temporal.
 */
function finalizeChunkedConversion_(payload) {
  const uploadId = String((payload && payload.uploadId) || '');
  const cache = CacheService.getScriptCache();
  const rawSession = cache.get('DOC2PDF_UP_' + uploadId);
  if (!rawSession) {
    throw new Error('No se encontró la sesión de subida para finalizar la conversión.');
  }

  const session = JSON.parse(rawSession);
  assertSessionOwnership_(session, 'finalización');

  const config = getConfig_();
  const sheetOptsAudit = session.sheetOptions ? formatSheetOptionsForAudit_(session.sheetOptions) : '';

  try {
    if (!session.googleFileId) {
      throw new Error('La subida aún no ha finalizado en Google Drive.');
    }

    const pdfFileName = buildPdfFileName_(session.fileName);
    const pdfBlob = exportGoogleFileToPdfBlob_(
      session.googleFileId,
      session.exportType,
      pdfFileName,
      session.sheetOptions
    );
    const transfer = preparePdfTransfer_(
      pdfBlob,
      pdfFileName,
      session.jobId,
      session.userEmail,
      session.tempFolderId
    );
    const durationMs = Date.now() - session.startedAtMs;

    incrementUserDailyUsage_(session.userEmail);
    logAuditJob_({
      jobId: session.jobId,
      userEmail: session.displayEmail,
      origin: 'Local',
      fileName: session.fileName,
      ext: session.ext,
      family: session.family,
      inputBytes: session.totalBytes,
      pdfBytes: transfer.pdfSizeBytes,
      durationMs: durationMs,
      status: 'OK',
      detail: 'Conversión fragmentada completada',
      sheetOptions: sheetOptsAudit
    }, config);

    return Object.assign({
      jobId: session.jobId,
      ext: session.ext,
      family: session.family,
      inputSizeBytes: session.totalBytes,
      durationMs: durationMs
    }, transfer);
  } catch (err) {
    const durationMs = Date.now() - session.startedAtMs;
    logAuditJob_({
      jobId: session.jobId,
      userEmail: session.displayEmail,
      origin: 'Local',
      fileName: session.fileName,
      ext: session.ext,
      family: session.family,
      inputBytes: session.totalBytes,
      pdfBytes: 0,
      durationMs: durationMs,
      status: 'ERROR',
      detail: err.message,
      sheetOptions: sheetOptsAudit
    }, config);
    throw err;
  } finally {
    if (session.googleFileId) {
      removeDriveFilePermanently_(session.googleFileId);
    }
    cache.remove('DOC2PDF_UP_' + uploadId);
  }
}

/**
 * Cancela y limpia una sesión de subida fragmentada tras validar propiedad.
 */
function abortChunkedUpload_(payload) {
  const uploadId = String((payload && payload.uploadId) || '');
  if (!uploadId) return { cleaned: false };
  const cache = CacheService.getScriptCache();
  const rawSession = cache.get('DOC2PDF_UP_' + uploadId);
  if (rawSession) {
    const session = JSON.parse(rawSession);
    assertSessionOwnership_(session, 'cancelación');
    if (session.googleFileId) {
      removeDriveFilePermanently_(session.googleFileId);
    }
    cache.remove('DOC2PDF_UP_' + uploadId);
  }
  return { cleaned: true };
}

/**
 * Descarga un bloque parcial de un PDF temporal grande (>7 MB) usando cabecera Range.
 */
function fetchPdfDownloadChunk_(payload) {
  const token = String((payload && payload.downloadToken) || '');
  const chunkIndex = Number(payload && payload.chunkIndex);
  const raw = CacheService.getScriptCache().get('DOC2PDF_DL_' + token);
  if (!raw) {
    throw new Error('El enlace temporal de descarga ha expirado.');
  }
  const dlSession = JSON.parse(raw);
  assertSessionOwnership_(dlSession, 'descarga');

  if (isNaN(chunkIndex) || chunkIndex < 0 || chunkIndex >= dlSession.totalChunks) {
    throw new Error('Índice de fragmento de descarga fuera de rango.');
  }

  const start = chunkIndex * DOWNLOAD_CHUNK_BYTES;
  const end = Math.min(dlSession.totalBytes - 1, start + DOWNLOAD_CHUNK_BYTES - 1);

  const resp = UrlFetchApp.fetch(
    'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(dlSession.fileId) + '?alt=media&supportsAllDrives=true',
    {
      method: 'get',
      headers: {
        Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
        Range: 'bytes=' + start + '-' + end
      },
      muteHttpExceptions: true
    }
  );

  const code = resp.getResponseCode();
  if (code !== 200 && code !== 206) {
    throw new Error('No se pudo leer el fragmento del PDF temporal (HTTP ' + code + ').');
  }

  return {
    chunkBase64: Utilities.base64Encode(resp.getContent()),
    chunkIndex: chunkIndex,
    totalChunks: dlSession.totalChunks
  };
}

/**
 * Elimina inmediatamente el PDF temporal de salida una vez descargado por el propietario de sesión.
 */
function completePdfDownload_(payload) {
  const token = String((payload && payload.downloadToken) || '');
  if (!token) return { cleaned: false };
  const cache = CacheService.getScriptCache();
  const raw = cache.get('DOC2PDF_DL_' + token);
  if (raw) {
    const dlSession = JSON.parse(raw);
    assertSessionOwnership_(dlSession, 'cierre de descarga');
    removeDriveFilePermanently_(dlSession.fileId);
    cache.remove('DOC2PDF_DL_' + token);
    return { cleaned: true };
  }
  return { cleaned: false };
}

/**
 * Convierte a PDF un archivo accesible mediante enlace compartido de Google Drive,
 * verificando primero contra Confused Deputy que esté compartido con el usuario o su dominio.
 */
function convertDriveUrl_(payload) {
  const startedAt = Date.now();
  const config = getConfig_();
  const jobId = generateJobId_();
  const rawUrl = String((payload && payload.url) || '').trim();
  const userCtx = assertAuthorizedForJob_(config, 'Enlace Drive', rawUrl || '(URL)', '', '');

  let tempGoogleId = '';
  let resolvedName = rawUrl;
  let resolvedExt = '';
  let resolvedFamily = '';
  let inputSizeBytes = 0;
  let sheetOptsAudit = '';

  try {
    if (!config.PERMITIR_ENLACE_DRIVE) {
      throw new Error('La conversión mediante enlace de Google Drive está deshabilitada por administración.');
    }

    const parsedId = extractDriveFileId_(rawUrl);
    if (!parsedId) {
      throw new Error('No se ha reconocido un enlace o ID válido de Google Drive / Docs / Sheets / Slides.');
    }

    const fieldsBase = 'id, name, mimeType, size, trashed, parents, owners(emailAddress), permissions(type, emailAddress, domain, role)';
    let meta = null;
    try {
      meta = Drive.Files.get(parsedId, {
        fields: fieldsBase + ', shortcutDetails(targetId)',
        supportsAllDrives: true
      });
    } catch (driveErr) {
      throw new Error('No es posible acceder al archivo. Asegúrate de compartirlo en modo lectura con el dominio del centro o con la cuenta del servicio.');
    }

    if (meta && meta.mimeType === 'application/vnd.google-apps.shortcut' && meta.shortcutDetails && meta.shortcutDetails.targetId) {
      meta = Drive.Files.get(meta.shortcutDetails.targetId, {
        fields: fieldsBase,
        supportsAllDrives: true
      });
    }

    if (!meta || meta.trashed) {
      throw new Error('El archivo indicado no existe o está en la papelera.');
    }

    const tempFolderId = getOrCreateTempFolder_(config);
    assertDriveFileSharedWithRequester_(meta, userCtx, tempFolderId);

    const actualFileId = meta.id;
    resolvedName = meta.name || actualFileId;
    inputSizeBytes = Number(meta.size) || 0;

    // Caso A: Documento nativo de Google Workspace (Docs, Sheets, Slides)
    const nativeFormat = getNativeFormatByMime_(meta.mimeType, config);
    if (nativeFormat) {
      resolvedExt = nativeFormat.ext;
      resolvedFamily = nativeFormat.family;
      const sheetOpts = nativeFormat.exportType === 'spreadsheets'
        ? normalizeSheetOptions_(payload && payload.sheetOptions)
        : null;
      sheetOptsAudit = sheetOpts ? formatSheetOptionsForAudit_(sheetOpts) : '';

      const pdfFileName = buildPdfFileName_(resolvedName);
      const pdfBlob = exportGoogleFileToPdfBlob_(actualFileId, nativeFormat.exportType, pdfFileName, sheetOpts);
      const transfer = preparePdfTransfer_(pdfBlob, pdfFileName, jobId, userCtx.email, tempFolderId);
      const durationMs = Date.now() - startedAt;

      incrementUserDailyUsage_(userCtx.email);
      logAuditJob_({
        jobId: jobId,
        userEmail: userCtx.displayEmail,
        origin: 'Enlace Drive',
        fileName: resolvedName,
        ext: resolvedExt,
        family: resolvedFamily,
        inputBytes: inputSizeBytes,
        pdfBytes: transfer.pdfSizeBytes,
        durationMs: durationMs,
        status: 'OK',
        detail: 'Exportado directamente desde documento nativo Workspace',
        sheetOptions: sheetOptsAudit
      }, config);

      return Object.assign({
        jobId: jobId,
        sourceName: resolvedName,
        ext: resolvedExt,
        family: resolvedFamily,
        inputSizeBytes: inputSizeBytes,
        durationMs: durationMs
      }, transfer);
    }

    // Caso B: Archivo de escritorio almacenado en Google Drive
    resolvedExt = extractExtension_(resolvedName);
    let desktopFormat = getFormatByExtension_(resolvedExt, config);
    if (!desktopFormat) {
      desktopFormat = getFormatByMimeType_(meta.mimeType, config);
      if (desktopFormat && !resolvedExt) {
        resolvedExt = desktopFormat.ext;
      }
    }

    if (!desktopFormat) {
      throw new Error('Tipo de archivo no compatible o familia deshabilitada (' + (meta.mimeType || resolvedExt || 'desconocido') + ').');
    }

    resolvedFamily = desktopFormat.family;
    const maxBytes = config.TAMANO_MAX_MB * 1024 * 1024;
    if (inputSizeBytes > maxBytes) {
      throw new Error('El archivo en Drive supera el límite de ' + config.TAMANO_MAX_MB + ' MB.');
    }

    const sheetOpts = desktopFormat.exportType === 'spreadsheets'
      ? normalizeSheetOptions_(payload && payload.sheetOptions)
      : null;
    sheetOptsAudit = sheetOpts ? formatSheetOptionsForAudit_(sheetOpts) : '';

    try {
      const copied = Drive.Files.copy(
        {
          name: '__tmp_' + jobId,
          mimeType: desktopFormat.targetGoogleMime,
          parents: [tempFolderId]
        },
        actualFileId,
        { fields: 'id', supportsAllDrives: true }
      );
      tempGoogleId = copied.id;
    } catch (copyErr) {
      // Fallback: descargar el blob binario y crear el documento temporal convertido
      const sourceBlob = DriveApp.getFileById(actualFileId).getBlob().setContentType(desktopFormat.sourceMime);
      const created = Drive.Files.create(
        {
          name: '__tmp_' + jobId,
          mimeType: desktopFormat.targetGoogleMime,
          parents: [tempFolderId]
        },
        sourceBlob,
        { fields: 'id', supportsAllDrives: true }
      );
      tempGoogleId = created.id;
    }

    const pdfFileName = buildPdfFileName_(resolvedName);
    const pdfBlob = exportGoogleFileToPdfBlob_(tempGoogleId, desktopFormat.exportType, pdfFileName, sheetOpts);
    const transfer = preparePdfTransfer_(pdfBlob, pdfFileName, jobId, userCtx.email, tempFolderId);
    const durationMs = Date.now() - startedAt;

    incrementUserDailyUsage_(userCtx.email);
    logAuditJob_({
      jobId: jobId,
      userEmail: userCtx.displayEmail,
      origin: 'Enlace Drive',
      fileName: resolvedName,
      ext: resolvedExt,
      family: resolvedFamily,
      inputBytes: inputSizeBytes,
      pdfBytes: transfer.pdfSizeBytes,
      durationMs: durationMs,
      status: 'OK',
      detail: 'Conversión desde archivo en Drive completada',
      sheetOptions: sheetOptsAudit
    }, config);

    return Object.assign({
      jobId: jobId,
      sourceName: resolvedName,
      ext: resolvedExt,
      family: resolvedFamily,
      inputSizeBytes: inputSizeBytes,
      durationMs: durationMs
    }, transfer);
  } catch (err) {
    const durationMs = Date.now() - startedAt;
    const isSecurityRejection = err && (err.code === 'DRIVE_ACCESS_FORBIDDEN' || err.code === 'DRIVE_ACCESS_NOT_SHARED');
    logAuditJob_({
      jobId: jobId,
      userEmail: userCtx.displayEmail,
      origin: 'Enlace Drive',
      fileName: resolvedName,
      ext: resolvedExt,
      family: resolvedFamily,
      inputBytes: inputSizeBytes,
      pdfBytes: 0,
      durationMs: durationMs,
      status: isSecurityRejection ? 'RECHAZADO' : 'ERROR',
      detail: isSecurityRejection ? ('[' + err.code + '] ' + err.message) : err.message,
      sheetOptions: sheetOptsAudit
    }, config);
    throw err;
  } finally {
    if (tempGoogleId) {
      removeDriveFilePermanently_(tempGoogleId);
    }
  }
}
