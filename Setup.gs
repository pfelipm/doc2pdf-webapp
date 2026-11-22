/**
 * Setup.gs
 * Menú personalizado en Google Sheets, inicialización de pestañas ("Ajustes" y "Auditoría")
 * y controladores de los diálogos modales de administración protegidos por rol y dialogToken.
 */

/**
 * Añade el menú personalizado de Doc2PDF al abrir la hoja de cálculo.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const maintenanceMenu = ui.createMenu('🛠️ Mantenimiento')
    .addItem('🔄 Limpiar caché de ajustes', 'menuClearConfigCache')
    .addItem('🧹 Purgar archivos temporales ahora', 'menuPurgeTempFiles')
    .addItem('⏱️ Instalar trigger diario de limpieza', 'menuInstallMaintenanceTrigger')
    .addItem('🔍 Verificar formatos Drive API', 'menuCheckDriveImportFormats');

  ui.createMenu('📄 Doc2PDF')
    .addItem('🌐 Abrir WebApp', 'openWebAppDialog')
    .addItem('🚀 Configurar URL de WebApp', 'showDeployDialog')
    .addSeparator()
    .addItem('📊 Resumen y analítica', 'showDashboardDialog')
    .addSeparator()
    .addItem('✨ Inicializar / reparar hoja', 'setupSheet')
    .addSubMenu(maintenanceMenu)
    .addSeparator()
    .addItem('ℹ️ Acerca de Doc2PDF', 'showAboutDialog')
    .addToUi();
}

/**
 * Inicializa o pide confirmación si la hoja ya contiene Ajustes configurados.
 */
function setupSheet() {
  const adminEmail = assertSpreadsheetAdmin_('setupSheet');
  const ss = getSpreadsheet_();
  const configSheet = ss.getSheetByName(SHEET_CONFIG_NAME);
  if (!configSheet || configSheet.getLastRow() < 2) {
    executeSetupSheet_({ resetAudit: false });
    SpreadsheetApp.getUi().alert(
      'Estructura inicializada',
      'Se han creado las pestañas "Ajustes" y "Auditoría" junto con la carpeta temporal de Drive.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  const template = HtmlService.createTemplateFromFile('InitDialog');
  template.APP_NAME = APP_NAME;
  template.accentColor = getConfig_().COLOR_ACENTO;
  template.dialogToken = createAdminDialogToken_(adminEmail);
  const html = template.evaluate().setWidth(510).setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

/**
 * Crea o repara las pestañas "Ajustes" y "Auditoría" conservando la URL de WebApp y,
 * salvo que se solicite explícitamente, el historial existente en "Auditoría".
 * Privada con sufijo "_" para impedir su invocación directa desde google.script.run.
 */
function executeSetupSheet_(options) {
  const opts = options || {};
  const ss = getSpreadsheet_();
  if (!ss) {
    throw new Error('No hay una hoja de cálculo activa.');
  }

  // 1. Pestaña Ajustes
  let configSheet = ss.getSheetByName(SHEET_CONFIG_NAME);
  if (!configSheet) {
    configSheet = ss.insertSheet(SHEET_CONFIG_NAME, 0);
  }

  const rows = [['Ajuste', 'Valor', 'Descripción']];
  for (let i = 0; i < DEFAULT_CONFIG_DEFS.length; i++) {
    rows.push([
      DEFAULT_CONFIG_DEFS[i][0],
      DEFAULT_CONFIG_DEFS[i][1],
      DEFAULT_CONFIG_DEFS[i][2]
    ]);
  }
  const usedRows = rows.length;
  const usedCols = 3;

  if (configSheet.getMaxRows() < usedRows) {
    configSheet.insertRowsAfter(configSheet.getMaxRows(), usedRows - configSheet.getMaxRows());
  }
  if (configSheet.getMaxColumns() < usedCols) {
    configSheet.insertColumnsAfter(configSheet.getMaxColumns(), usedCols - configSheet.getMaxColumns());
  }

  configSheet.clear();
  configSheet.getRange(1, 1, configSheet.getMaxRows(), configSheet.getMaxColumns()).clearDataValidations();
  configSheet.getRange(1, 1, usedRows, usedCols).setValues(rows);

  // Estilo de cabecera y filas cebra suaves
  configSheet.getRange('A1:C1')
    .setFontWeight('bold')
    .setBackground('#0f2b5b')
    .setFontColor('#ffffff')
    .setVerticalAlignment('middle');
  configSheet.setFrozenRows(1);

  for (let r = 2; r <= usedRows; r++) {
    const bg = r % 2 === 0 ? '#ffffff' : '#f8fafc';
    configSheet.getRange(r, 1, 1, 3).setBackground(bg).setVerticalAlignment('middle');
    configSheet.getRange(r, 1).setFontWeight('bold').setFontColor('#1e293b');
    configSheet.getRange(r, 3).setFontColor('#475569');

    const def = DEFAULT_CONFIG_DEFS[r - 2];
    const key = def[0];
    const valType = def[3];
    const valueCell = configSheet.getRange(r, 2);

    if (valType === 'checkbox') {
      valueCell.setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireCheckbox('SI', 'NO')
          .setAllowInvalid(false)
          .build()
      );
    } else if (valType === 'audit_mode') {
      valueCell.setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInList([AUDIT_MODE_COMPLETO, AUDIT_MODE_ANONIMIZADO, AUDIT_MODE_NINGUNO], true)
          .setAllowInvalid(false)
          .setHelpText('Selecciona COMPLETO, ANONIMIZADO o NINGUNO.')
          .build()
      );
    } else if (valType === 'number') {
      const numSpec = CONFIG_NUMERIC_LIMITS[key];
      if (numSpec) {
        valueCell.setDataValidation(
          SpreadsheetApp.newDataValidation()
            .requireNumberBetween(numSpec.min, numSpec.max)
            .setAllowInvalid(false)
            .setHelpText(numSpec.help)
            .build()
        );
      }
    }
  }

  configSheet.setColumnWidth(1, 235);
  configSheet.setColumnWidth(2, 340);
  configSheet.getRange(2, 2, usedRows - 1, 1).setWrap(true);
  configSheet.getRange(2, 3, usedRows - 1, 1).setWrap(false);
  SpreadsheetApp.flush();
  configSheet.autoResizeColumn(3);
  configSheet.setColumnWidth(3, Math.round(configSheet.getColumnWidth(3) * 1.08) + 32);

  SpreadsheetApp.flush();
  configSheet.setRowHeight(1, 32);
  configSheet.autoResizeRows(2, usedRows - 1);
  for (let r = 2; r <= usedRows; r++) {
    if (configSheet.getRowHeight(r) < 28) {
      configSheet.setRowHeight(r, 28);
    }
  }

  configSheet.getRange(1, 1, usedRows, usedCols).setBorder(
    true, true, true, true, true, true, '#e2e8f0', SpreadsheetApp.BorderStyle.SOLID
  );

  // Recortar filas y columnas sobrantes fuera de la tabla de Ajustes
  if (configSheet.getMaxRows() > usedRows) {
    configSheet.deleteRows(usedRows + 1, configSheet.getMaxRows() - usedRows);
  }
  if (configSheet.getMaxColumns() > usedCols) {
    configSheet.deleteColumns(usedCols + 1, configSheet.getMaxColumns() - usedCols);
  }

  // 2. Pestaña Auditoría
  let auditSheet = ss.getSheetByName(SHEET_AUDIT_NAME);
  if (!auditSheet) {
    auditSheet = ss.insertSheet(SHEET_AUDIT_NAME, 1);
  }

  const auditCols = AUDIT_HEADERS.length;
  const existingFilter = auditSheet.getFilter();
  if (existingFilter) {
    existingFilter.remove();
  }

  if (opts.resetAudit === true) {
    auditSheet.clear();
  }

  if (auditSheet.getMaxRows() < 2) {
    auditSheet.insertRowsAfter(auditSheet.getMaxRows(), 2 - auditSheet.getMaxRows());
  }
  if (auditSheet.getMaxColumns() < auditCols) {
    auditSheet.insertColumnsAfter(auditSheet.getMaxColumns(), auditCols - auditSheet.getMaxColumns());
  }

  auditSheet.getRange(1, 1, 1, auditCols).setValues([AUDIT_HEADERS]);

  styleAuditHeader_(auditSheet);
  if (auditSheet.getLastRow() >= 2) {
    auditSheet.getRange(2, 1, auditSheet.getLastRow() - 1, 1).setNumberFormat(getRegionalDateTimeFormat_(ss));
    clearColumnLinksAsPlainText_(auditSheet, 2, auditSheet.getLastRow() - 1, 5);
  }

  const targetAuditRows = Math.max(2, auditSheet.getLastRow());
  const targetAuditCols = Math.max(auditCols, auditSheet.getLastColumn());
  if (auditSheet.getMaxRows() > targetAuditRows) {
    auditSheet.deleteRows(targetAuditRows + 1, auditSheet.getMaxRows() - targetAuditRows);
  }
  if (auditSheet.getMaxColumns() > targetAuditCols) {
    auditSheet.deleteColumns(targetAuditCols + 1, auditSheet.getMaxColumns() - targetAuditCols);
  }

  auditSheet.getRange(1, 1, targetAuditRows, auditCols).createFilter();

  // Eliminar la pestaña genérica "Hoja 1" / "Sheet1" solo si está totalmente vacía
  const defaultSheetNames = ['Hoja 1', 'Sheet1'];
  for (let i = 0; i < defaultSheetNames.length; i++) {
    const candidate = ss.getSheetByName(defaultSheetNames[i]);
    if (candidate && ss.getSheets().length > 2 && candidate.getLastRow() === 0 && candidate.getLastColumn() === 0) {
      ss.deleteSheet(candidate);
    }
  }

  PropertiesService.getScriptProperties().deleteProperty(PROP_TEMP_FOLDER_ID);
  clearConfigCache_();
  const cfg = getConfig_(true);
  getOrCreateTempFolder_(cfg);

  return { success: true };
}

/**
 * Escapa texto para interpolación segura en cadenas HTML de servidor.
 */
function escapeHtmlServer_(val) {
  return String(val === null || val === undefined ? '' : val)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/**
 * Abre la WebApp en una nueva pestaña o muestra el asistente de despliegue si no hay URL guardada.
 */
function openWebAppDialog() {
  assertSpreadsheetAdmin_('openWebAppDialog');
  const url = getSavedWebAppUrl_();
  if (!url) {
    showDeployDialog();
    return;
  }

  const cfg = getConfig_();
  const safeUrl = escapeHtmlServer_(url);
  const safeJsUrl = JSON.stringify(url);
  const safeAppName = escapeHtmlServer_(cfg.NOMBRE_APP);
  const safeAccent = sanitizeHexColor_(cfg.COLOR_ACENTO, '#1a73e8');
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html lang="es">
      <head>
        <base target="_top">
        <meta charset="utf-8">
        <style>
          * { box-sizing: border-box; }
          body {
            margin: 0;
            font-family: system-ui, -apple-system, "Segoe UI", Roboto, sans-serif;
            background: #f8fafc;
            color: #0f172a;
          }
          .hdr {
            background: ${safeAccent};
            color: #fff;
            padding: 14px 20px;
            display: flex;
            align-items: center;
            gap: 10px;
            font-weight: 600;
            font-size: 16px;
          }
          .body {
            padding: 18px 20px;
            text-align: center;
          }
          .msg {
            margin: 0 0 16px 0;
            font-size: 13.5px;
            color: #334155;
            line-height: 1.45;
          }
          .actions {
            display: flex;
            justify-content: center;
            gap: 10px;
          }
          .btn {
            display: inline-flex;
            align-items: center;
            gap: 6px;
            border: 0;
            border-radius: 8px;
            padding: 9px 16px;
            font-size: 13.5px;
            font-weight: 600;
            cursor: pointer;
            text-decoration: none;
          }
          .btn-primary { background: ${safeAccent}; color: #fff; }
          .btn-secondary { background: #e2e8f0; color: #1e293b; }
          .tip { margin-top: 12px; font-size: 12px; color: #64748b; }
        </style>
      </head>
      <body>
        <div class="hdr">
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg>
          <span>Abrir ${safeAppName}</span>
        </div>
        <div class="body">
          <p class="msg">Si tu navegador bloqueó la ventana emergente automática, utiliza los botones siguientes:</p>
          <div class="actions">
            <a class="btn btn-primary" href="${safeUrl}" target="_blank" rel="noopener" onclick="setTimeout(function(){ google.script.host.close(); }, 250)">Abrir WebApp</a>
            <button type="button" class="btn btn-secondary" onclick="copyLink(this)">Copiar enlace</button>
          </div>
          <div id="tip" class="tip">Permite ventanas emergentes para abrir la app directamente la próxima vez.</div>
        </div>
        <script>
          var TARGET_URL = ${safeJsUrl};
          function copyLink(btn) {
            navigator.clipboard.writeText(TARGET_URL).then(function() {
              btn.textContent = '¡Enlace copiado!';
            });
          }
          try {
            var w = window.open(TARGET_URL, '_blank');
            if (w && !w.closed && typeof w.closed !== 'undefined') {
              google.script.host.close();
            }
          } catch (e) {}
        </script>
      </body>
    </html>
  `).setWidth(460).setHeight(225);

  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

/**
 * Abre el asistente modal para configurar y guardar la URL de la WebApp.
 */
function showDeployDialog() {
  const adminEmail = assertSpreadsheetAdmin_('showDeployDialog');
  const cfg = getConfig_();
  const template = HtmlService.createTemplateFromFile('DeployDialog');
  template.currentUrl = getSavedWebAppUrl_() || '';
  template.appName = cfg.NOMBRE_APP;
  template.accentColor = cfg.COLOR_ACENTO;
  template.editorUrl = 'https://script.google.com/home/projects/' + ScriptApp.getScriptId() + '/edit';
  template.dialogToken = createAdminDialogToken_(adminEmail);
  const html = template.evaluate().setWidth(620).setHeight(675);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

/**
 * Abre el diálogo HTML rico y vistoso de Resumen y analítica (sustituye a la pestaña de fórmulas).
 */
function showDashboardDialog() {
  const adminEmail = assertSpreadsheetAdmin_('showDashboardDialog');
  const cfg = getConfig_();
  const template = HtmlService.createTemplateFromFile('DashboardDialog');
  template.appName = cfg.NOMBRE_APP;
  template.accentColor = cfg.COLOR_ACENTO;
  template.version = APP_VERSION;
  template.locale = getSpreadsheetLocaleBcp47_(getSpreadsheet_());
  template.dialogToken = createAdminDialogToken_(adminEmail);
  const html = template.evaluate().setWidth(1040).setHeight(960);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

/**
 * Abre el modal de atribución "Acerca de Doc2PDF".
 */
function showAboutDialog() {
  assertSpreadsheetAdmin_('showAboutDialog');
  const cfg = getConfig_();
  const template = HtmlService.createTemplateFromFile('AboutDialog');
  template.VERSION = APP_VERSION;
  template.APP_NAME = APP_NAME;
  template.APP_REPO_URL = APP_REPO_URL;
  template.APP_AUTHOR_NAME = APP_AUTHOR_NAME;
  template.APP_AUTHOR_HANDLE = APP_AUTHOR_HANDLE;
  template.APP_AUTHOR_URL = APP_AUTHOR_URL;
  template.accentColor = cfg.COLOR_ACENTO;
  const html = template.evaluate().setWidth(490).setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

/**
 * Acción de menú: invalida la caché de ajustes.
 */
function menuClearConfigCache() {
  assertSpreadsheetAdmin_('menuClearConfigCache');
  clearConfigCache_();
  getSpreadsheet_().toast('Caché de ajustes limpiada. La próxima petición leerá la hoja directamente.', APP_NAME, 4);
}

/**
 * Acción de menú: purga manualmente todos los archivos de la carpeta temporal.
 */
function menuPurgeTempFiles() {
  assertSpreadsheetAdmin_('menuPurgeTempFiles');
  const res = purgeOrphanTempFiles_(0);
  SpreadsheetApp.getUi().alert(
    'Purga de temporales completada',
    'Archivos inspeccionados: ' + res.scanned + '\nArchivos eliminados permanentemente: ' + res.deleted + '\nErrores: ' + res.errors,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Acción de menú: instala el trigger diario de mantenimiento.
 */
function menuInstallMaintenanceTrigger() {
  assertSpreadsheetAdmin_('menuInstallMaintenanceTrigger');
  installMaintenanceTrigger_();
  SpreadsheetApp.getUi().alert(
    'Trigger instalado',
    'Se ha programado el mantenimiento diario automático (03:00–04:00 Europe/Madrid) para purgar temporales huérfanos y aplicar la retención de Auditoría.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Acción de menú: consulta Drive.About.get() para verificar formatos soportados en el tenant actual.
 */
function menuCheckDriveImportFormats() {
  assertSpreadsheetAdmin_('menuCheckDriveImportFormats');
  const about = Drive.About.get({ fields: 'importFormats, maxImportSizes' });
  const importFormats = (about && about.importFormats) || {};
  const normalizedImportFormats = {};
  Object.keys(importFormats).forEach(function (mimeKey) {
    normalizedImportFormats[String(mimeKey).toLowerCase()] = importFormats[mimeKey];
  });
  const cfg = getConfig_();
  const enabled = getEnabledFormats_(cfg);
  const lines = [];
  let verified = 0;

  for (let i = 0; i < enabled.length; i++) {
    const fmt = enabled[i];
    const targets = normalizedImportFormats[String(fmt.sourceMime || '').toLowerCase()] || [];
    const ok = targets.indexOf(fmt.targetGoogleMime) !== -1;
    if (ok) verified++;
    lines.push((ok ? '✅ ' : '⚠️ ') + '.' + fmt.ext + ' (' + fmt.sourceMime + ')');
  }

  SpreadsheetApp.getUi().alert(
    'Verificación Drive API v3 (' + verified + '/' + enabled.length + ' confirmados)',
    lines.join('\n'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/* ============================================================================
 * Endpoints RPC expuestos a los diálogos HTML (protegidos por rol + dialogToken)
 * ========================================================================== */

function apiGetDashboardAnalytics(daysFilter, dialogToken) {
  try {
    assertAdminDialogCall_(dialogToken, 'apiGetDashboardAnalytics');
    const data = getDashboardAnalytics_(daysFilter);
    return { success: true, data: data };
  } catch (e) {
    return { success: false, code: e.code || 'ADMIN_ERROR', error: e.message };
  }
}

function apiPurgeTempFilesFromDialog(dialogToken) {
  try {
    assertAdminDialogCall_(dialogToken, 'apiPurgeTempFilesFromDialog');
    const res = purgeOrphanTempFiles_(0);
    return { success: true, result: res };
  } catch (e) {
    return { success: false, code: e.code || 'ADMIN_ERROR', error: e.message };
  }
}

function apiInstallTriggerFromDialog(dialogToken) {
  try {
    assertAdminDialogCall_(dialogToken, 'apiInstallTriggerFromDialog');
    const res = installMaintenanceTrigger_();
    return { success: true, result: res };
  } catch (e) {
    return { success: false, code: e.code || 'ADMIN_ERROR', error: e.message };
  }
}

function apiExecuteSetupSheet(options, dialogToken) {
  try {
    assertAdminDialogCall_(dialogToken, 'apiExecuteSetupSheet');
    executeSetupSheet_(options);
    return { success: true };
  } catch (e) {
    return { success: false, code: e.code || 'ADMIN_ERROR', error: e.message };
  }
}

function apiSaveWebAppUrl(url, dialogToken) {
  try {
    assertAdminDialogCall_(dialogToken, 'apiSaveWebAppUrl');
    return saveWebAppUrl_(url);
  } catch (e) {
    return { success: false, code: e.code || 'ADMIN_ERROR', error: e.message };
  }
}
