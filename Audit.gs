/**
 * Audit.gs
 * Registro concurrente seguro en la pestaña "Auditoría", protección contra
 * inyección de fórmulas, anonimización opcional, retención y analítica para Dashboard.
 */

const AUDIT_HEADERS = [
  'Marca temporal',
  'ID trabajo',
  'Usuario',
  'Origen',
  'Archivo',
  'Extensión',
  'Familia',
  'Tamaño Entrada (KB)',
  'Tamaño PDF (KB)',
  'Duración (ms)',
  'Estado',
  'Detalle',
  'Opciones PDF'
];

/**
 * Protege valores de texto contra Formula Injection en hojas de cálculo y acota longitud.
 */
function sanitizeCellValue_(val) {
  if (val === null || val === undefined) return '';
  let str = String(val).replace(/[\r\n]+/g, ' ').trim();
  if (str.length > 500) {
    str = str.slice(0, 497) + '...';
  }
  if (/^[=+\-@\t\r]/.test(str)) {
    str = "'" + str;
  }
  return str;
}

/**
 * Calcula un hash SHA-256 corto (10 caracteres hex) en minúsculas.
 */
function shortSha256_(text) {
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(text || ''),
    Utilities.Charset.UTF_8
  );
  let hex = '';
  for (let i = 0; i < digest.length && hex.length < 10; i++) {
    const byteVal = digest[i] < 0 ? digest[i] + 256 : digest[i];
    const part = byteVal.toString(16);
    hex += part.length === 1 ? '0' + part : part;
  }
  return hex.slice(0, 10);
}

/**
 * Formatea el nombre del archivo según la política REGISTRAR_NOMBRE_ARCHIVO.
 */
function formatAuditedFileName_(rawName, ext, mode) {
  const cleanExt = ext ? String(ext).toLowerCase() : '';
  const policy = String(mode || AUDIT_MODE_COMPLETO).trim().toUpperCase();
  if (policy === AUDIT_MODE_NINGUNO) {
    return cleanExt ? '(omitido).' + cleanExt : '(omitido)';
  }
  if (policy === AUDIT_MODE_ANONIMIZADO) {
    return 'doc_' + shortSha256_(rawName || 'file') + (cleanExt ? '.' + cleanExt : '');
  }
  return sanitizeCellValue_(rawName || '(sin nombre)');
}

/**
 * Asegura que exista la pestaña Auditoría con su cabecera estilizada.
 */
function ensureAuditSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_AUDIT_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_AUDIT_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, AUDIT_HEADERS.length).setValues([AUDIT_HEADERS]);
    styleAuditHeader_(sheet);
  }
  return sheet;
}

/**
 * Aplica estilo a la cabecera de la pestaña Auditoría.
 */
function styleAuditHeader_(sheet) {
  const auditCols = AUDIT_HEADERS.length;
  if (sheet.getMaxRows() < 2) {
    sheet.insertRowsAfter(sheet.getMaxRows(), 2 - sheet.getMaxRows());
  }
  if (sheet.getMaxColumns() < auditCols) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), auditCols - sheet.getMaxColumns());
  }

  const headerRange = sheet.getRange(1, 1, 1, auditCols);
  headerRange
    .setFontWeight('bold')
    .setBackground('#0f2b5b')
    .setFontColor('#ffffff')
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('left')
    .setWrap(true);
  sheet.setRowHeight(1, 34);
  sheet.setFrozenRows(1);

  const widths = [170, 185, 220, 108, 240, 118, 136, 196, 168, 148, 110, 260, 200];
  for (let i = 0; i < widths.length; i++) {
    sheet.setColumnWidth(i + 1, widths[i]);
  }
}

/**
 * Registra un trabajo en la pestaña "Auditoría" usando ScriptLock solo durante la escritura.
 */
function logAuditJob_(entry, config) {
  const cfg = config || getConfig_();
  const ss = getSpreadsheet_();
  if (!ss) return false;

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = ensureAuditSheet_(ss);
    const inputKB = entry.inputBytes ? Math.round((entry.inputBytes / 1024) * 10) / 10 : 0;
    const pdfKB = entry.pdfBytes ? Math.round((entry.pdfBytes / 1024) * 10) / 10 : 0;
    const durationMs = typeof entry.durationMs === 'number' ? Math.max(0, Math.round(entry.durationMs)) : 0;

    const row = [
      new Date(),
      sanitizeCellValue_(entry.jobId || ''),
      sanitizeCellValue_(entry.userEmail || 'usuario-dominio'),
      sanitizeCellValue_(entry.origin || 'Local'),
      formatAuditedFileName_(entry.fileName, entry.ext, cfg.REGISTRAR_NOMBRE_ARCHIVO),
      sanitizeCellValue_(entry.ext || ''),
      sanitizeCellValue_(entry.family || ''),
      inputKB,
      pdfKB,
      durationMs,
      sanitizeCellValue_(entry.status || 'OK'),
      sanitizeCellValue_(entry.detail || ''),
      sanitizeCellValue_(entry.sheetOptions || '')
    ];

    sheet.appendRow(row);
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      sheet.getRange(lastRow, 1).setNumberFormat(getRegionalDateTimeFormat_(ss));
      setPlainTextNoLink_(sheet.getRange(lastRow, 5), row[4]);
      sheet.getRange(lastRow, 8, 1, 2).setNumberFormat('#,##0.0');
      sheet.getRange(lastRow, 10).setNumberFormat('#,##0');
    }
    return true;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Elimina filas de Auditoría más antiguas que retentionDays.
 * @return {{removed: number, kept: number}}
 */
function purgeOldAuditRows_(retentionDays) {
  const days = parseInt(retentionDays, 10);
  if (isNaN(days) || days <= 0) {
    return { removed: 0, kept: 0 };
  }

  const ss = getSpreadsheet_();
  if (!ss) return { removed: 0, kept: 0 };
  const sheet = ss.getSheetByName(SHEET_AUDIT_NAME);
  if (!sheet || sheet.getLastRow() <= 1) {
    return { removed: 0, kept: 0 };
  }

  const cutoffMs = Date.now() - days * 24 * 60 * 60 * 1000;
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const lastRow = sheet.getLastRow();
    const values = sheet.getRange(2, 1, lastRow - 1, AUDIT_HEADERS.length).getValues();
    const keptRows = [];
    let removed = 0;

    for (let i = 0; i < values.length; i++) {
      const ts = values[i][0];
      const tsMs = ts instanceof Date ? ts.getTime() : Date.parse(ts);
      if (!tsMs || tsMs >= cutoffMs) {
        keptRows.push(values[i]);
      } else {
        removed++;
      }
    }

    if (removed > 0) {
      sheet.getRange(2, 1, lastRow - 1, AUDIT_HEADERS.length).clearContent();
      if (keptRows.length > 0) {
        sheet.getRange(2, 1, keptRows.length, AUDIT_HEADERS.length).setValues(keptRows);
        sheet.getRange(2, 1, keptRows.length, 1).setNumberFormat(getRegionalDateTimeFormat_(ss));
        clearColumnLinksAsPlainText_(sheet, 2, keptRows.length, 5);
        sheet.getRange(2, 8, keptRows.length, 2).setNumberFormat('#,##0.0');
        sheet.getRange(2, 10, keptRows.length, 1).setNumberFormat('#,##0');
      }
    }
    return { removed: removed, kept: keptRows.length };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Devuelve la zona horaria configurada en la hoja de cálculo contenedora
 * (con respaldo a la zona del script y Europe/Madrid).
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet=} ssOpt
 * @return {string}
 */
function getLocalTimeZone_(ssOpt) {
  try {
    const ss = ssOpt || getSpreadsheet_();
    if (ss && typeof ss.getSpreadsheetTimeZone === 'function') {
      const tz = String(ss.getSpreadsheetTimeZone() || '').trim();
      if (tz) return tz;
    }
  } catch (e) {}
  try {
    if (typeof Session !== 'undefined' && Session && typeof Session.getScriptTimeZone === 'function') {
      const scriptTz = String(Session.getScriptTimeZone() || '').trim();
      if (scriptTz) return scriptTz;
    }
  } catch (e) {}
  return 'Europe/Madrid';
}

/**
 * Devuelve el formato de fecha/hora acorde con la configuración regional de la hoja.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet=} ssOpt
 * @return {string}
 */
function getRegionalDateTimeFormat_(ssOpt) {
  let locale = '';
  try {
    const ss = ssOpt || getSpreadsheet_();
    if (ss && typeof ss.getSpreadsheetLocale === 'function') {
      locale = String(ss.getSpreadsheetLocale() || '').trim().toLowerCase();
    }
  } catch (e) {}

  if (locale === 'en_us' || locale === 'en_ph' || locale === 'en_bz') {
    return 'mm/dd/yyyy hh:mm:ss';
  }
  if (/^(ja|zh|ko|hu|sv|lt)(_|$)/.test(locale)) {
    return 'yyyy/mm/dd hh:mm:ss';
  }
  return 'dd/mm/yyyy hh:mm:ss';
}

/**
 * Devuelve el identificador BCP 47 de configuración regional de la hoja (p. ej. 'es-ES').
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet=} ssOpt
 * @return {string}
 */
function getSpreadsheetLocaleBcp47_(ssOpt) {
  try {
    const ss = ssOpt || getSpreadsheet_();
    if (ss && typeof ss.getSpreadsheetLocale === 'function') {
      const raw = String(ss.getSpreadsheetLocale() || '').trim();
      if (raw) {
        return raw.replace(/_/g, '-');
      }
    }
  } catch (e) {}
  return 'es-ES';
}

/**
 * Escribe texto plano en una celda desactivando la conversión automática de enlaces (p. ej. *.md).
 * @param {GoogleAppsScript.Spreadsheet.Range} range
 * @param {*} text
 */
function setPlainTextNoLink_(range, text) {
  if (!range) return;
  const str = text === null || text === undefined ? '' : String(text);
  range.setNumberFormat('@');
  if (typeof SpreadsheetApp !== 'undefined' && typeof SpreadsheetApp.newRichTextValue === 'function' && typeof range.setRichTextValue === 'function') {
    const builder = SpreadsheetApp.newRichTextValue().setText(str);
    if (str) {
      builder.setLinkUrl(null);
    }
    range.setRichTextValue(builder.build());
  }
}

/**
 * Limpia hipervínculos automáticos en una columna completa y la fuerza a formato texto plano ('@').
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} startRow
 * @param {number} numRows
 * @param {number} colIndex
 */
function clearColumnLinksAsPlainText_(sheet, startRow, numRows, colIndex) {
  if (!sheet || numRows <= 0) return;
  const range = sheet.getRange(startRow, colIndex, numRows, 1);
  range.setNumberFormat('@');
  if (typeof SpreadsheetApp === 'undefined' || typeof SpreadsheetApp.newRichTextValue !== 'function' || typeof range.setRichTextValues !== 'function') {
    return;
  }
  const values = range.getValues();
  const richValues = [];
  for (let i = 0; i < values.length; i++) {
    const str = values[i][0] === null || values[i][0] === undefined ? '' : String(values[i][0]);
    const builder = SpreadsheetApp.newRichTextValue().setText(str);
    if (str) {
      builder.setLinkUrl(null);
    }
    richValues.push([builder.build()]);
  }
  range.setRichTextValues(richValues);
}

/**
 * Formatea una fecha a yyyy-MM-dd usando la zona horaria local de la hoja (compatible con Node en tests).
 */
function formatIsoDay_(dateObj) {
  if (typeof Utilities !== 'undefined' && Utilities.formatDate) {
    return Utilities.formatDate(dateObj, getLocalTimeZone_(), 'yyyy-MM-dd');
  }
  const y = dateObj.getFullYear();
  const m = String(dateObj.getMonth() + 1).padStart(2, '0');
  const d = String(dateObj.getDate()).padStart(2, '0');
  return y + '-' + m + '-' + d;
}

/**
 * Formatea una marca temporal legible en zona local para la tabla del Dashboard.
 */
function formatTimestampDisplay_(dateObj) {
  if (typeof Utilities !== 'undefined' && Utilities.formatDate) {
    return Utilities.formatDate(dateObj, getLocalTimeZone_(), 'dd/MM HH:mm');
  }
  const dd = String(dateObj.getDate()).padStart(2, '0');
  const mm = String(dateObj.getMonth() + 1).padStart(2, '0');
  const hh = String(dateObj.getHours()).padStart(2, '0');
  const min = String(dateObj.getMinutes()).padStart(2, '0');
  return dd + '/' + mm + ' ' + hh + ':' + min;
}

/**
 * Lógica pura de agregación analítica sobre las filas de Auditoría para el Dashboard.
 * @param {Array<Array>} rows Filas de datos (sin cabecera).
 * @param {number} daysFilter 7, 30, 90 o 0 (todo el histórico).
 * @param {Date=} nowDate Fecha de referencia.
 * @param {Object=} healthEstado Estado operativo del servicio.
 */
function computeDashboardAnalytics_(rows, daysFilter, nowDate, healthEstado) {
  const now = nowDate || new Date();
  const nowMs = now.getTime();
  const filterDays = typeof daysFilter === 'number' && daysFilter > 0 ? daysFilter : 0;
  const cutoffMs = filterDays > 0 ? nowMs - filterDays * 24 * 60 * 60 * 1000 : 0;
  const todayIso = formatIsoDay_(now);

  let total = 0;
  let ok = 0;
  let error = 0;
  let rejected = 0;
  let todayCount = 0;
  let totalInputKB = 0;
  let totalPdfKB = 0;
  const okDurations = [];

  const usersSet = {};
  const userStatsMap = {};
  const familyCounts = {
    DOCUMENTOS: 0,
    HOJAS: 0,
    PRESENTACIONES: 0,
    TEXTO: 0,
    OTROS: 0
  };
  const extCounts = {};
  const originCounts = { local: 0, drive: 0 };
  const dailyMap = {};
  const recentJobs = [];

  const sourceRows = Array.isArray(rows) ? rows : [];
  const parsedRows = [];

  for (let i = 0; i < sourceRows.length; i++) {
    const r = sourceRows[i];
    if (!r || r.length < 11) continue;
    const rawTs = r[0];
    if (!rawTs) continue;
    const tsDate = rawTs instanceof Date ? rawTs : new Date(rawTs);
    const tsMs = tsDate.getTime();
    if (isNaN(tsMs)) continue;

    if (cutoffMs > 0 && tsMs < cutoffMs) continue;

    const status = String(r[10] || '').trim().toUpperCase();
    const user = String(r[2] || 'usuario-dominio').trim();
    const originRaw = String(r[3] || 'Local').trim();
    const file = String(r[4] || '');
    const ext = String(r[5] || '').trim().toLowerCase();
    const familyRaw = String(r[6] || '').trim().toUpperCase();
    const inputKB = Number(r[7]) || 0;
    const pdfKB = Number(r[8]) || 0;
    const durationMs = Number(r[9]) || 0;
    const detail = String(r[11] || '');
    const jobId = String(r[1] || '');
    const isoDay = formatIsoDay_(tsDate);

    parsedRows.push({
      tsDate: tsDate,
      tsMs: tsMs,
      isoDay: isoDay,
      jobId: jobId,
      user: user,
      origin: originRaw,
      file: file,
      ext: ext,
      family: familyRaw,
      inputKB: inputKB,
      pdfKB: pdfKB,
      durationMs: durationMs,
      status: status,
      detail: detail
    });
  }

  for (let i = 0; i < parsedRows.length; i++) {
    const item = parsedRows[i];
    total++;

    if (item.isoDay === todayIso && item.status === 'OK') {
      todayCount++;
    }

    if (!dailyMap[item.isoDay]) {
      dailyMap[item.isoDay] = {
        isoDate: item.isoDay,
        label: item.isoDay.slice(8, 10) + '/' + item.isoDay.slice(5, 7),
        ok: 0,
        error: 0,
        rejected: 0,
        total: 0
      };
    }
    dailyMap[item.isoDay].total++;

    if (item.status === 'OK') {
      ok++;
      dailyMap[item.isoDay].ok++;
      totalInputKB += item.inputKB;
      totalPdfKB += item.pdfKB;
      if (item.durationMs > 0) okDurations.push(item.durationMs);
    } else if (item.status === 'RECHAZADO') {
      rejected++;
      dailyMap[item.isoDay].rejected++;
    } else {
      error++;
      dailyMap[item.isoDay].error++;
    }

    usersSet[item.user] = true;
    if (!userStatsMap[item.user]) {
      userStatsMap[item.user] = { email: item.user, total: 0, ok: 0, error: 0 };
    }
    userStatsMap[item.user].total++;
    if (item.status === 'OK') {
      userStatsMap[item.user].ok++;
    } else {
      userStatsMap[item.user].error++;
    }

    if (Object.prototype.hasOwnProperty.call(familyCounts, item.family)) {
      familyCounts[item.family]++;
    } else if (item.family) {
      familyCounts.OTROS++;
    }

    if (item.ext) {
      extCounts[item.ext] = (extCounts[item.ext] || 0) + 1;
    }

    if (item.origin.toLowerCase().indexOf('drive') !== -1) {
      originCounts.drive++;
    } else {
      originCounts.local++;
    }
  }

  // Rellenar días faltantes si filterDays está entre 1 y 30 para una serie temporal continua
  if (filterDays > 0 && filterDays <= 30) {
    for (let d = filterDays - 1; d >= 0; d--) {
      const dayDate = new Date(nowMs - d * 24 * 60 * 60 * 1000);
      const iso = formatIsoDay_(dayDate);
      if (!dailyMap[iso]) {
        dailyMap[iso] = {
          isoDate: iso,
          label: iso.slice(8, 10) + '/' + iso.slice(5, 7),
          ok: 0,
          error: 0,
          rejected: 0,
          total: 0
        };
      }
    }
  }

  let dailySeries = Object.keys(dailyMap)
    .sort()
    .map(function (k) { return dailyMap[k]; });

  // Si hay más de 30 puntos (p. ej. filtro de 90 días o histórico), agrupar por bloques o conservar últimos 30 días activos
  if (dailySeries.length > 30) {
    const bucketSize = Math.ceil(dailySeries.length / 24);
    const grouped = [];
    for (let i = 0; i < dailySeries.length; i += bucketSize) {
      const slice = dailySeries.slice(i, i + bucketSize);
      const agg = {
        isoDate: slice[0].isoDate,
        label: slice.length > 1
          ? slice[0].label + '–' + slice[slice.length - 1].label
          : slice[0].label,
        ok: 0,
        error: 0,
        rejected: 0,
        total: 0
      };
      for (let j = 0; j < slice.length; j++) {
        agg.ok += slice[j].ok;
        agg.error += slice[j].error;
        agg.rejected += slice[j].rejected;
        agg.total += slice[j].total;
      }
      grouped.push(agg);
    }
    dailySeries = grouped;
  }

  // Duraciones media y p95
  let avgDurationMs = 0;
  let p95DurationMs = 0;
  if (okDurations.length > 0) {
    const sumDur = okDurations.reduce(function (a, b) { return a + b; }, 0);
    avgDurationMs = Math.round(sumDur / okDurations.length);
    const sortedDur = okDurations.slice().sort(function (a, b) { return a - b; });
    const p95Idx = Math.min(sortedDur.length - 1, Math.ceil(sortedDur.length * 0.95) - 1);
    p95DurationMs = Math.round(sortedDur[p95Idx]);
  }

  // Distribución por familia
  const familyTotal = Object.keys(familyCounts).reduce(function (acc, key) {
    return acc + familyCounts[key];
  }, 0);

  const familyPalette = {
    DOCUMENTOS: { label: 'Documentos', color: '#1a73e8' },
    HOJAS: { label: 'Hojas de cálculo', color: '#188038' },
    PRESENTACIONES: { label: 'Presentaciones', color: '#e37400' },
    TEXTO: { label: 'Texto y marcado', color: '#7c3aed' },
    OTROS: { label: 'Otros', color: '#64748b' }
  };

  const byFamily = Object.keys(familyCounts)
    .filter(function (fam) { return familyCounts[fam] > 0; })
    .map(function (fam) {
      const cnt = familyCounts[fam];
      return {
        family: fam,
        label: familyPalette[fam].label,
        color: familyPalette[fam].color,
        count: cnt,
        percent: familyTotal > 0 ? Math.round((cnt / familyTotal) * 1000) / 10 : 0
      };
    })
    .sort(function (a, b) { return b.count - a.count; });

  // Top 8 extensiones
  const extTotal = Object.keys(extCounts).reduce(function (acc, k) { return acc + extCounts[k]; }, 0);
  const byExtension = Object.keys(extCounts)
    .map(function (ext) {
      const cnt = extCounts[ext];
      return {
        ext: ext,
        count: cnt,
        percent: extTotal > 0 ? Math.round((cnt / extTotal) * 1000) / 10 : 0
      };
    })
    .sort(function (a, b) { return b.count - a.count; })
    .slice(0, 8);

  // Top 8 usuarios
  const topUsers = Object.keys(userStatsMap)
    .map(function (email) {
      const u = userStatsMap[email];
      return {
        email: u.email,
        total: u.total,
        ok: u.ok,
        error: u.error,
        successRate: u.total > 0 ? Math.round((u.ok / u.total) * 1000) / 10 : 0
      };
    })
    .sort(function (a, b) { return b.total - a.total; })
    .slice(0, 8);

  // Últimos 20 trabajos ordenados del más reciente al más antiguo
  const sortedRecent = parsedRows
    .slice()
    .sort(function (a, b) { return b.tsMs - a.tsMs; })
    .slice(0, 20);

  for (let i = 0; i < sortedRecent.length; i++) {
    const job = sortedRecent[i];
    recentJobs.push({
      timestamp: formatTimestampDisplay_(job.tsDate),
      jobId: job.jobId,
      user: job.user,
      origin: job.origin,
      file: job.file,
      ext: job.ext,
      family: job.family,
      inputKB: job.inputKB,
      pdfKB: job.pdfKB,
      durationMs: job.durationMs,
      status: job.status,
      detail: job.detail
    });
  }

  return {
    daysFilter: filterDays,
    generatedAt: formatTimestampDisplay_(now),
    kpis: {
      total: total,
      ok: ok,
      error: error,
      rejected: rejected,
      successRate: total > 0 ? Math.round((ok / total) * 1000) / 10 : 0,
      todayCount: todayCount,
      uniqueUsers: Object.keys(usersSet).length,
      totalInputMB: Math.round((totalInputKB / 1024) * 100) / 100,
      totalPdfMB: Math.round((totalPdfKB / 1024) * 100) / 100,
      avgDurationMs: avgDurationMs,
      p95DurationMs: p95DurationMs
    },
    dailySeries: dailySeries,
    byFamily: byFamily,
    byExtension: byExtension,
    byOrigin: originCounts,
    topUsers: topUsers,
    recentJobs: recentJobs,
    health: healthEstado || {}
  };
}

/**
 * Lee Auditoría, obtiene el estado del sistema y devuelve la analítica para el Dashboard.
 * @param {number=} daysFilter 7, 30, 90 o 0 (todo).
 * @return {Object}
 */
function getDashboardAnalytics_(daysFilter) {
  const cfg = getConfig_(true);
  const ss = getSpreadsheet_();
  const sheet = ss ? ss.getSheetByName(SHEET_AUDIT_NAME) : null;

  let rows = [];
  let totalAuditRows = 0;
  if (sheet && sheet.getLastRow() >= 2) {
    totalAuditRows = sheet.getLastRow() - 1;
    rows = sheet.getRange(2, 1, totalAuditRows, AUDIT_HEADERS.length).getValues();
  }

  const health = {
    serviceActive: cfg.SERVICIO_ACTIVO,
    tempFilesCount: countTempFiles_(cfg),
    hasTrigger: hasMaintenanceTrigger_(),
    maxSizeMB: cfg.TAMANO_MAX_MB,
    dailyLimit: cfg.LIMITE_DIARIO_USUARIO,
    auditNameMode: cfg.REGISTRAR_NOMBRE_ARCHIVO,
    webAppConfigured: Boolean(getSavedWebAppUrl_()),
    totalAuditRows: totalAuditRows
  };

  const parsedFilter = typeof daysFilter === 'number' ? daysFilter : parseInt(daysFilter, 10);
  const safeFilter = [0, 7, 30, 90].indexOf(parsedFilter) !== -1 ? parsedFilter : 30;
  return computeDashboardAnalytics_(rows, safeFilter, new Date(), health);
}
