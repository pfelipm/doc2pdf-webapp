/**
 * @OnlyCurrentDoc
 *
 * WebApp.gs
 * Punto de entrada HTTP de la WebApp (doGet), inclusión privada de parciales HTML
 * y endpoints RPC expuestos al cliente mediante google.script.run.
 */

/**
 * Sirve la aplicación web Doc2PDF.
 */
function doGet() {
  const config = getConfig_();
  const template = HtmlService.createTemplateFromFile('Index');
  template.appTitle = config.NOMBRE_APP || APP_NAME;
  template.accentColor = config.COLOR_ACENTO || '#1a73e8';
  template.appVersion = APP_VERSION;

  return template.evaluate()
    .setTitle(template.appTitle)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Incluye parciales HTML (Styles.html, ClientScript.html) dentro de Index.html.
 * Privada con sufijo "_" para no exponerla a google.script.run.
 */
function include_(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Envuelve errores para devolver un objeto RPC homogéneo y seguro.
 */
function buildRpcError_(err) {
  return {
    success: false,
    code: (err && err.code) || 'CONVERSION_ERROR',
    error: (err && err.message) ? err.message : String(err || 'Error desconocido')
  };
}

/**
 * Devuelve la configuración pública, formatos habilitados y estado del usuario activo.
 */
function apiGetBootstrap() {
  try {
    const config = getConfig_();
    const userCtx = getUserContext_(config);
    return {
      success: true,
      data: getPublicClientConfig_(userCtx, config)
    };
  } catch (e) {
    return buildRpcError_(e);
  }
}

/**
 * Convierte un archivo local (<= 6 MB) enviado directamente en Base64.
 */
function apiConvertLocalDirect(payload) {
  try {
    return {
      success: true,
      data: convertLocalFileDirect_(payload)
    };
  } catch (e) {
    return buildRpcError_(e);
  }
}

/**
 * Inicia una subida fragmentada para archivos locales > 6 MB.
 */
function apiInitChunkedUpload(payload) {
  try {
    return {
      success: true,
      data: initChunkedUpload_(payload)
    };
  } catch (e) {
    return buildRpcError_(e);
  }
}

/**
 * Envía un fragmento de archivo a la sesión Resumable Upload activa.
 */
function apiUploadChunk(payload) {
  try {
    return {
      success: true,
      data: uploadFileChunk_(payload)
    };
  } catch (e) {
    return buildRpcError_(e);
  }
}

/**
 * Finaliza una subida fragmentada y exporta el documento a PDF.
 */
function apiFinalizeChunkedUpload(payload) {
  try {
    return {
      success: true,
      data: finalizeChunkedConversion_(payload)
    };
  } catch (e) {
    return buildRpcError_(e);
  }
}

/**
 * Aborta y limpia una subida fragmentada cancelada o fallida.
 */
function apiAbortChunkedUpload(payload) {
  try {
    return {
      success: true,
      data: abortChunkedUpload_(payload)
    };
  } catch (e) {
    return buildRpcError_(e);
  }
}

/**
 * Convierte a PDF un documento compartido mediante URL o ID de Google Drive.
 */
function apiConvertDriveUrl(payload) {
  try {
    return {
      success: true,
      data: convertDriveUrl_(payload)
    };
  } catch (e) {
    return buildRpcError_(e);
  }
}

/**
 * Descarga un fragmento de un PDF temporal grande (> 7 MB).
 */
function apiFetchPdfChunk(payload) {
  try {
    return {
      success: true,
      data: fetchPdfDownloadChunk_(payload)
    };
  } catch (e) {
    return buildRpcError_(e);
  }
}

/**
 * Confirma la finalización de descarga de un PDF grande y lo borra permanentemente de Drive.
 */
function apiCompletePdfDownload(payload) {
  try {
    return {
      success: true,
      data: completePdfDownload_(payload)
    };
  } catch (e) {
    return buildRpcError_(e);
  }
}
