/**
 * Cleanup.gs
 * Gestión de la carpeta temporal de Drive, borrado permanente de temporales
 * y tareas de mantenimiento programadas.
 */

/**
 * Obtiene el ID de la carpeta padre donde reside la hoja de cálculo de control,
 * o cadena vacía si no se puede determinar.
 * @return {string}
 */
function getSpreadsheetParentFolderId_() {
  try {
    const ss = getSpreadsheet_();
    const ssId = ss && typeof ss.getId === 'function' ? ss.getId() : '';
    if (!ssId) return '';
    const meta = Drive.Files.get(ssId, {
      fields: 'parents',
      supportsAllDrives: true
    });
    if (meta && Array.isArray(meta.parents) && meta.parents.length > 0) {
      return String(meta.parents[0] || '').trim();
    }
  } catch (e) {
    console.warn('No se pudo obtener la carpeta padre de la hoja de control:', e.message);
  }
  return '';
}

/**
 * Obtiene o crea la carpeta temporal dedicada dentro de la carpeta donde reside la hoja
 * de control (con respaldo a la raíz de Mi unidad si no se puede determinar padre).
 * @param {Object=} config Configuración actual opcional.
 * @return {string} ID de la carpeta temporal.
 */
function getOrCreateTempFolder_(config) {
  const cfg = config || getConfig_();
  const props = PropertiesService.getScriptProperties();
  const candidates = [];

  if (cfg.CARPETA_TEMP_ID) {
    candidates.push(cfg.CARPETA_TEMP_ID);
  }
  const propFolderId = props.getProperty(PROP_TEMP_FOLDER_ID);
  if (propFolderId && candidates.indexOf(propFolderId) === -1) {
    candidates.push(propFolderId);
  }

  for (let i = 0; i < candidates.length; i++) {
    const id = candidates[i];
    try {
      const meta = Drive.Files.get(id, {
        fields: 'id, name, mimeType, trashed',
        supportsAllDrives: true
      });
      if (meta && !meta.trashed && meta.mimeType === 'application/vnd.google-apps.folder') {
        if (propFolderId !== meta.id) {
          props.setProperty(PROP_TEMP_FOLDER_ID, meta.id);
        }
        if (!cfg.CARPETA_TEMP_ID) {
          setConfigValue_('CARPETA_TEMP_ID', meta.id);
        }
        return meta.id;
      }
    } catch (e) {
      console.warn('Carpeta temporal candidata no válida (' + id + '):', e.message);
    }
  }

  const sheetParentId = getSpreadsheetParentFolderId_();

  // Buscar por nombre en la carpeta de la hoja (o raíz como respaldo) para evitar duplicados
  try {
    const parentClause = sheetParentId ? "'" + sheetParentId + "' in parents" : "'root' in parents";
    const existing = Drive.Files.list({
      q: parentClause + " and name = '" + TEMP_FOLDER_DEFAULT_NAME + "' and mimeType = 'application/vnd.google-apps.folder' and trashed = false",
      fields: 'files(id, name)',
      pageSize: 5,
      supportsAllDrives: true,
      includeItemsFromAllDrives: true
    });
    if (existing && existing.files && existing.files.length > 0) {
      const foundId = existing.files[0].id;
      props.setProperty(PROP_TEMP_FOLDER_ID, foundId);
      if (!cfg.CARPETA_TEMP_ID) {
        setConfigValue_('CARPETA_TEMP_ID', foundId);
      }
      return foundId;
    }
  } catch (e) {
    console.warn('No se pudo buscar carpeta temporal existente:', e.message);
  }

  const folderResource = {
    name: TEMP_FOLDER_DEFAULT_NAME,
    mimeType: 'application/vnd.google-apps.folder',
    description: 'Carpeta temporal automática de Doc2PDF. Sus archivos se purgan automáticamente.'
  };
  if (sheetParentId) {
    folderResource.parents = [sheetParentId];
  }

  const created = Drive.Files.create(
    folderResource,
    null,
    { fields: 'id', supportsAllDrives: true }
  );

  props.setProperty(PROP_TEMP_FOLDER_ID, created.id);
  if (!cfg.CARPETA_TEMP_ID) {
    setConfigValue_('CARPETA_TEMP_ID', created.id);
  }
  return created.id;
}

/**
 * Elimina permanentemente un archivo temporal de Drive sin pasarlo por la papelera.
 * Si falla el borrado permanente, intenta enviarlo a la papelera como respaldo.
 * @param {string} fileId
 * @return {boolean}
 */
function removeDriveFilePermanently_(fileId) {
  if (!fileId) return false;
  try {
    Drive.Files.remove(fileId, { supportsAllDrives: true });
    return true;
  } catch (removeErr) {
    const msg = String(removeErr && removeErr.message ? removeErr.message : removeErr);
    if (msg.indexOf('404') !== -1 || msg.indexOf('File not found') !== -1) {
      return true;
    }
    console.warn('Drive.Files.remove falló para ' + fileId + '; intentando fallback a papelera:', msg);
    try {
      DriveApp.getFileById(fileId).setTrashed(true);
      return true;
    } catch (trashErr) {
      console.error('No se pudo eliminar temporal ' + fileId + ':', trashErr.message);
      return false;
    }
  }
}

/**
 * Cuenta los archivos activos actualmente dentro de la carpeta temporal.
 * @param {Object=} config
 * @return {number}
 */
function countTempFiles_(config) {
  try {
    const folderId = getOrCreateTempFolder_(config);
    let total = 0;
    let pageToken = null;
    do {
      const res = Drive.Files.list({
        q: "'" + folderId + "' in parents and trashed = false",
        fields: 'nextPageToken, files(id)',
        pageSize: 100,
        pageToken: pageToken,
        supportsAllDrives: true,
        includeItemsFromAllDrives: true
      });
      if (res && res.files) {
        total += res.files.length;
      }
      pageToken = res ? res.nextPageToken : null;
    } while (pageToken);
    return total;
  } catch (e) {
    console.warn('Error contando temporales:', e.message);
    return 0;
  }
}

/**
 * Purga permanentemente archivos huérfanos de la carpeta temporal que superen
 * la antigüedad indicada en minutos.
 * @param {number=} maxAgeMinutes Antigüedad mínima en minutos (30 por defecto; 0 = purgar todos).
 * @return {{scanned: number, deleted: number, errors: number}}
 */
function purgeOrphanTempFiles_(maxAgeMinutes) {
  const ageLimit = typeof maxAgeMinutes === 'number' && maxAgeMinutes >= 0 ? maxAgeMinutes : 30;
  const cfg = getConfig_();
  const folderId = getOrCreateTempFolder_(cfg);
  const ss = getSpreadsheet_();
  const ssId = ss && typeof ss.getId === 'function' ? ss.getId() : '';
  const cutoffMs = Date.now() - ageLimit * 60 * 1000;

  let scanned = 0;
  let deleted = 0;
  let errors = 0;
  let pageToken = null;

  do {
    const res = Drive.Files.list({
      q: "'" + folderId + "' in parents and trashed = false",
      fields: 'nextPageToken, files(id, name, mimeType, createdTime, modifiedTime)',
      pageSize: 100,
      pageToken: pageToken,
      supportsAllDrives: true,
      includeItemsFromAllDrives: true
    });

    const files = res && res.files ? res.files : [];
    for (let i = 0; i < files.length; i++) {
      const f = files[i];
      if (!f || !f.id || (ssId && f.id === ssId) || f.mimeType === 'application/vnd.google-apps.folder') {
        continue;
      }
      scanned++;
      const refIso = f.modifiedTime || f.createdTime;
      const refMs = refIso ? Date.parse(refIso) : 0;
      if (ageLimit === 0 || !refMs || refMs <= cutoffMs) {
        if (removeDriveFilePermanently_(f.id)) {
          deleted++;
        } else {
          errors++;
        }
      }
    }
    pageToken = res ? res.nextPageToken : null;
  } while (pageToken);

  return { scanned: scanned, deleted: deleted, errors: errors };
}

/**
 * Comprueba si el trigger diario de mantenimiento está instalado.
 * @return {boolean}
 */
function hasMaintenanceTrigger_() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    for (let i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === 'runDailyMaintenance') {
        return true;
      }
    }
    return false;
  } catch (e) {
    console.warn('No se pudieron inspeccionar los triggers:', e.message);
    return false;
  }
}

/**
 * Instala (reemplazando previos) el trigger diario de mantenimiento a las 03:00.
 * @return {{installed: boolean, handler: string}}
 */
function installMaintenanceTrigger_() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'runDailyMaintenance') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger('runDailyMaintenance')
    .timeBased()
    .everyDays(1)
    .atHour(3)
    .create();

  return { installed: true, handler: 'runDailyMaintenance' };
}

/**
 * Rutina ejecutada por el trigger diario: purga temporales huérfanos (>30 min)
 * y aplica la política de retención en Auditoría.
 * Protegida frente a invocaciones no autorizadas vía google.script.run.
 */
function runDailyMaintenance(e) {
  assertAdminOrTimeTrigger_(e);
  const cfg = getConfig_(true);
  const tempRes = purgeOrphanTempFiles_(30);
  const auditRes = purgeOldAuditRows_(cfg.RETENCION_AUDITORIA_DIAS);
  console.log('Mantenimiento diario completado:', JSON.stringify({
    temp: tempRes,
    audit: auditRes
  }));
  return { temp: tempRes, audit: auditRes };
}
