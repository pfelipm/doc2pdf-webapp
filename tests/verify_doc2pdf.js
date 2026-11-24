const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');
const crypto = require('crypto');

const srcDir = path.resolve(__dirname, '..');
const rootDir = srcDir;

// 1. Validación JSON y @OnlyCurrentDoc
console.log('1) Validando manifiesto, scopes y @OnlyCurrentDoc...');
for (const jsonPath of [
  path.join(srcDir, '.clasp.json.example'),
  path.join(srcDir, 'appsscript.json')
]) {
  JSON.parse(fs.readFileSync(jsonPath, 'utf8'));
  console.log('   OK JSON:', path.relative(rootDir, jsonPath));
}
const claspExample = JSON.parse(fs.readFileSync(path.join(srcDir, '.clasp.json.example'), 'utf8'));
assert.strictEqual(claspExample.rootDir, '.', '.clasp.json.example debe apuntar a rootDir "."');

const appsscript = JSON.parse(fs.readFileSync(path.join(srcDir, 'appsscript.json'), 'utf8'));
assert.strictEqual(appsscript.webapp.executeAs, 'USER_DEPLOYING');
assert.strictEqual(appsscript.webapp.access, 'DOMAIN');
assert(appsscript.oauthScopes.includes('https://www.googleapis.com/auth/spreadsheets.currentonly'), 'Debe usar spreadsheets.currentonly');
assert(appsscript.oauthScopes.includes('https://www.googleapis.com/auth/groups'), 'Debe incluir scope groups para verificación directa de Google Groups');
assert(!appsscript.oauthScopes.includes('https://www.googleapis.com/auth/spreadsheets'), 'No debe usar scope amplio spreadsheets');
assert(!appsscript.oauthScopes.some((s) => s.includes('send_mail')), 'No debe incluir script.send_mail');
assert(!appsscript.oauthScopes.some((s) => s.includes('admin.directory')), 'No debe requerir Admin SDK Directory');

assert(fs.readFileSync(path.join(srcDir, 'Config.gs'), 'utf8').includes('@OnlyCurrentDoc'), 'Config.gs debe declarar @OnlyCurrentDoc');
assert(fs.readFileSync(path.join(srcDir, 'WebApp.gs'), 'utf8').includes('@OnlyCurrentDoc'), 'WebApp.gs debe declarar @OnlyCurrentDoc');

// 2. Validación sintáctica y verificación estricta de funciones públicas
console.log('2) Validando sintaxis JS y superficie pública (sufijo "_")...');
const gsFiles = fs.readdirSync(srcDir).filter((f) => f.endsWith('.gs')).sort();
const publicFunctions = [];

for (const file of gsFiles) {
  const code = fs.readFileSync(path.join(srcDir, file), 'utf8');
  new vm.Script(code, { filename: file });
  for (const m of code.matchAll(/^function\s+([A-Za-z0-9_]+)\s*\(/gm)) {
    const fnName = m[1];
    if (!fnName.endsWith('_')) {
      publicFunctions.push({ file, fnName });
    }
  }
  console.log('   OK .gs:', file);
}

const allowedPublicSet = new Set([
  // WebApp.gs
  'doGet', 'apiGetBootstrap', 'apiConvertLocalDirect', 'apiInitChunkedUpload',
  'apiUploadChunk', 'apiFinalizeChunkedUpload', 'apiAbortChunkedUpload',
  'apiConvertDriveUrl', 'apiFetchPdfChunk', 'apiCompletePdfDownload',
  // Setup.gs
  'onOpen', 'setupSheet', 'openWebAppDialog', 'showDeployDialog',
  'showDashboardDialog', 'showAboutDialog', 'menuClearConfigCache',
  'menuPurgeTempFiles', 'menuInstallMaintenanceTrigger', 'menuCheckDriveImportFormats',
  'menuCheckAccessLists',
  'apiGetDashboardAnalytics', 'apiGetAuditJobs', 'apiSaveAccessLists', 'apiPurgeTempFilesFromDialog',
  'apiInstallTriggerFromDialog', 'apiRemoveTriggerFromDialog',
  'apiExecuteSetupSheet', 'apiSaveWebAppUrl',
  // Cleanup.gs
  'runDailyMaintenance'
]);

const unexpectedPublic = publicFunctions.filter((p) => !allowedPublicSet.has(p.fnName));
assert.strictEqual(
  unexpectedPublic.length,
  0,
  'Funciones internas expuestas sin "_": ' + JSON.stringify(unexpectedPublic)
);
assert.strictEqual(publicFunctions.length, allowedPublicSet.size, 'Lista exacta de funciones públicas');
console.log(`   OK: exactamente ${publicFunctions.length} funciones públicas autorizadas (0 fugas internas).`);

const htmlFiles = fs.readdirSync(srcDir).filter((f) => f.endsWith('.html')).sort();
for (const file of htmlFiles) {
  const html = fs.readFileSync(path.join(srcDir, file), 'utf8');
  const scriptRegex = /<script\b[^>]*>([\s\S]*?)<\/script>/gi;
  let match;
  let blockIdx = 0;
  while ((match = scriptRegex.exec(html)) !== null) {
    blockIdx++;
    const sanitized = match[1].replace(/<\?!?=[\s\S]*?\?>/g, '"__TEMPLATE__"');
    new vm.Script(sanitized, { filename: `${file}#script${blockIdx}` });
  }
  console.log(`   OK .html (${blockIdx} bloques script):`, file);
}

// Comprobar cero alert()/confirm() nativos en diálogos ricos
const dashHtml = fs.readFileSync(path.join(srcDir, 'DashboardDialog.html'), 'utf8');
assert(!/\balert\s*\(/.test(dashHtml), 'DashboardDialog.html no debe usar alert() nativo');
for (const dlg of ['TriggerDialog.html', 'PurgeTempDialog.html', 'AccessCheckDialog.html']) {
  const txt = fs.readFileSync(path.join(srcDir, dlg), 'utf8');
  assert(!/\b(alert|confirm)\s*\(/.test(txt), `${dlg} no debe usar alert() ni confirm() nativos`);
}

// 3. Verificación de contratos RPC entre HTML y backend
console.log('3) Verificando contratos RPC y dialogToken...');
const backendCode = ['WebApp.gs', 'Setup.gs']
  .map((f) => fs.readFileSync(path.join(srcDir, f), 'utf8'))
  .join('\n');
const exportedFns = new Set(
  Array.from(backendCode.matchAll(/function\s+([a-zA-Z0-9_]+)\s*\(/g)).map((m) => m[1])
);

const rpcCalls = [];
for (const file of [
  'ClientScript.html', 'DashboardDialog.html', 'InitDialog.html',
  'DeployDialog.html', 'TriggerDialog.html', 'PurgeTempDialog.html',
  'AccessCheckDialog.html'
]) {
  const txt = fs.readFileSync(path.join(srcDir, file), 'utf8');
  for (const m of txt.matchAll(/runServer\(\s*['"]([a-zA-Z0-9_]+)['"]/g)) {
    rpcCalls.push({ file, fn: m[1] });
  }
  for (const m of txt.matchAll(/\.(api[A-Za-z0-9_]+)\s*\(/g)) {
    rpcCalls.push({ file, fn: m[1] });
  }
}
for (const call of rpcCalls) {
  assert(exportedFns.has(call.fn), `RPC inexistente ${call.fn} llamado desde ${call.file}`);
}

for (const dlg of [
  'DashboardDialog.html', 'InitDialog.html', 'DeployDialog.html',
  'TriggerDialog.html', 'PurgeTempDialog.html', 'AccessCheckDialog.html'
]) {
  const txt = fs.readFileSync(path.join(srcDir, dlg), 'utf8');
  assert(txt.includes("DIALOG_TOKEN = '<?= dialogToken ?>'"), `${dlg} debe inyectar DIALOG_TOKEN`);
  assert(txt.includes('DIALOG_TOKEN)'), `${dlg} debe pasar DIALOG_TOKEN en sus llamadas RPC`);
}
console.log(`   OK: ${rpcCalls.length} llamadas RPC y contratos dialogToken verificados.`);

// 4. Tests unitarios y de intrusión con mocks GAS en VM
console.log('4) Ejecutando pruebas funcionales y de seguridad (mocks GAS)...');

const memCache = new Map();
const memProps = new Map();
const auditRows = [];
const removedFiles = [];
const createdDriveFiles = [];
let mockListFiles = [];
const ajustesMetrics = {
  maxRows: 1000,
  maxCols: 26,
  colWidths: {},
  rowHeights: {},
  validations: {},
  wrapCalls: [],
  clearDataValidationsCount: 0,
  autoResizedCols: [],
  autoResizeRowsCalls: [],
  deletedRows: [],
  deletedCols: [],
  cellUpdates: {},
  singleUpdates: {}
};
const auditMetrics = {
  hasHeader: true,
  lastCol: 13,
  maxRows: 1000,
  maxCols: 26,
  colWidths: {},
  rowHeights: {},
  wrapCalls: [],
  numberFormats: [],
  richTextCalls: [],
  frozenRows: 0,
  filterRange: null,
  hasFilter: false,
  filterRemovedCount: 0,
  deletedRows: [],
  deletedCols: [],
  insertedRows: [],
  insertedCols: []
};
let shouldFailExport = false;
let serviceActiveSetting = 'SI';
let activeEmail = 'docente@centro.edu';
let uuidCounter = 0;
let customBlacklistValue = 'bloqueado@centro.edu';
let customWhitelistValue = '';
const uiAlerts = [];
const uiModals = [];
const mockTriggers = [];
const mockGroups = {};
let mockAboutResponse = {
  importFormats: {
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document': ['application/vnd.google-apps.document'],
    'application/msword': ['application/vnd.google-apps.document'],
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['application/vnd.google-apps.spreadsheet'],
    'application/vnd.openxmlformats-officedocument.presentationml.presentation': ['application/vnd.google-apps.presentation'],
    'text/plain': ['application/vnd.google-apps.document']
  },
  maxImportSizes: {
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document': String(50 * 1024 * 1024),
    'text/plain': String(10 * 1024 * 1024)
  }
};
let groupsLookupCount = 0;
let groupsHasUserCount = 0;

const defaultAjustesRows = () => [
  ['NOMBRE_APP', 'Doc2PDF Centro'],
  ['SUBTITULO_CENTRO', 'Microservicio'],
  ['COLOR_ACENTO', '#2563eb'],
  ['LOGO_URL', 'https://example.edu/logo.png'],
  ['SERVICIO_ACTIVO', serviceActiveSetting],
  ['MENSAJE_MANTENIMIENTO', 'En mantenimiento'],
  ['DOMINIOS_PERMITIDOS', 'centro.edu'],
  ['LISTA_NEGRA_USUARIOS', customBlacklistValue],
  ['LISTA_BLANCA_USUARIOS', customWhitelistValue],
  ['LIMITE_DIARIO_USUARIO', 30],
  ['TAMANO_MAX_MB', 20],
  ['MAX_ARCHIVOS_LOTE', 5],
  ['PERMITIR_DOCUMENTOS', 'SI'],
  ['PERMITIR_HOJAS', 'SI'],
  ['PERMITIR_PRESENTACIONES', 'SI'],
  ['PERMITIR_TEXTO', 'SI'],
  ['PERMITIR_ENLACE_DRIVE', 'SI'],
  ['VISOR_PDF_EMBEBIDO', 'NO'],
  ['REGISTRAR_NOMBRE_ARCHIVO', 'COMPLETO'],
  ['RETENCION_AUDITORIA_DIAS', 180],
  ['CARPETA_TEMP_ID', 'FOLDER_TEMP_123']
];

function makeRange(sheetName, row, col, numRows, numCols) {
  return {
    getValues() {
      if (sheetName === 'Ajustes') {
        const rows = defaultAjustesRows();
        if (ajustesMetrics.cellUpdates['22,2'] !== undefined) {
          rows[20][1] = ajustesMetrics.cellUpdates['22,2'];
        }
        if (ajustesMetrics.singleUpdates['9,2'] !== undefined) {
          rows[7][1] = ajustesMetrics.singleUpdates['9,2'];
        }
        if (ajustesMetrics.singleUpdates['10,2'] !== undefined) {
          rows[8][1] = ajustesMetrics.singleUpdates['10,2'];
        }
        return rows.slice(Math.max(0, row - 2), Math.max(0, row - 2) + (numRows || rows.length));
      }
      if (sheetName === 'Auditoría') {
        const startIdx = Math.max(0, (row || 2) - 2);
        const endIdx = numRows ? startIdx + numRows : auditRows.length;
        const sliced = auditRows.slice(startIdx, endIdx);
        if (col && numCols) {
          return sliced.map((r) => r.slice(col - 1, col - 1 + numCols));
        }
        return sliced;
      }
      return [];
    },
    setValue(v) {
      if (sheetName === 'Ajustes') {
        ajustesMetrics.cellUpdates[row + ',' + col] = v;
        ajustesMetrics.singleUpdates[row + ',' + col] = v;
      }
      return this;
    },
    setValues(vals) {
      if (sheetName === 'Ajustes' && Array.isArray(vals)) {
        for (let rIdx = 0; rIdx < vals.length; rIdx++) {
          for (let cIdx = 0; cIdx < vals[rIdx].length; cIdx++) {
            ajustesMetrics.cellUpdates[(row + rIdx) + ',' + (col + cIdx)] = vals[rIdx][cIdx];
          }
        }
      }
      if (sheetName === 'Auditoría' && row === 1) {
        auditMetrics.hasHeader = true;
        if (Array.isArray(vals) && vals[0]) {
          auditMetrics.lastCol = Math.max(auditMetrics.lastCol, col - 1 + vals[0].length);
        }
      }
      return this;
    },
    setNumberFormat(fmt) {
      if (sheetName === 'Auditoría') {
        auditMetrics.numberFormats.push({ row, col, numRows, fmt });
      }
      return this;
    },
    setRichTextValue(rt) {
      if (sheetName === 'Auditoría') {
        auditMetrics.richTextCalls.push({ row, col, numRows, values: [[rt]] });
      }
      return this;
    },
    setRichTextValues(rts) {
      if (sheetName === 'Auditoría') {
        auditMetrics.richTextCalls.push({ row, col, numRows, values: rts });
      }
      return this;
    },
    setFontWeight() { return this; },
    setBackground() { return this; },
    setFontColor() { return this; },
    setVerticalAlignment() { return this; },
    setHorizontalAlignment() { return this; },
    setWrap(wrap) {
      if (sheetName === 'Ajustes') {
        ajustesMetrics.wrapCalls.push({ row, col, numRows, numCols, wrap });
      }
      if (sheetName === 'Auditoría') {
        auditMetrics.wrapCalls.push({ row, col, numRows, numCols, wrap });
      }
      return this;
    },
    setDataValidation(rule) {
      if (sheetName === 'Ajustes') {
        ajustesMetrics.validations[row] = rule;
      }
      return this;
    },
    clearDataValidations() {
      if (sheetName === 'Ajustes') {
        ajustesMetrics.validations = {};
        ajustesMetrics.clearDataValidationsCount++;
      }
      return this;
    },
    setBorder() { return this; },
    createFilter() {
      if (sheetName === 'Auditoría') {
        auditMetrics.hasFilter = true;
        auditMetrics.filterRange = { row, col, numRows, numCols };
      }
      return this;
    },
    clear() { return this; },
    clearContent() { return this; },
    getFilter() { return true; }
  };
}

const sheetsMock = {
  Ajustes: {
    getLastRow: () => 22,
    getMaxRows: () => ajustesMetrics.maxRows,
    getMaxColumns: () => ajustesMetrics.maxCols,
    getRange: (r, c, nr, nc) => {
      if (typeof r === 'string') return makeRange('Ajustes', 1, 1, 1, 3);
      return makeRange('Ajustes', r, c, nr, nc);
    },
    clear() {},
    insertRowsAfter(pos, count) { ajustesMetrics.maxRows += count; },
    insertColumnsAfter(pos, count) { ajustesMetrics.maxCols += count; },
    deleteRows(pos, count) {
      ajustesMetrics.deletedRows.push({ pos, count });
      ajustesMetrics.maxRows -= count;
    },
    deleteColumns(pos, count) {
      ajustesMetrics.deletedCols.push({ pos, count });
      ajustesMetrics.maxCols -= count;
    },
    setRowHeight(r, h) { ajustesMetrics.rowHeights[r] = h; },
    getRowHeight(r) { return ajustesMetrics.rowHeights[r] || 21; },
    setFrozenRows() {},
    setColumnWidth(c, w) { ajustesMetrics.colWidths[c] = w; },
    getColumnWidth(c) { return ajustesMetrics.colWidths[c] || 100; },
    autoResizeColumn(c) {
      ajustesMetrics.autoResizedCols.push(c);
      ajustesMetrics.colWidths[c] = 560;
    },
    autoResizeRows(start, num) {
      ajustesMetrics.autoResizeRowsCalls.push({ start, num });
      for (let r = start; r < start + num; r++) {
        ajustesMetrics.rowHeights[r] = r === 14 ? 42 : 21;
      }
    }
  },
  'Auditoría': {
    getLastRow: () => (auditMetrics.hasHeader ? 1 + auditRows.length : 0),
    getLastColumn: () => auditMetrics.lastCol,
    getMaxRows: () => auditMetrics.maxRows,
    getMaxColumns: () => auditMetrics.maxCols,
    getRange: (r, c, nr, nc) => makeRange('Auditoría', r, c, nr, nc),
    appendRow: (row) => { auditMetrics.hasHeader = true; auditMetrics.lastCol = Math.max(auditMetrics.lastCol, row.length); auditRows.push(row); },
    clear() {
      auditRows.length = 0;
      auditMetrics.hasHeader = false;
      auditMetrics.lastCol = 0;
    },
    insertRowsAfter(pos, count) {
      auditMetrics.insertedRows.push({ pos, count });
      auditMetrics.maxRows += count;
    },
    insertColumnsAfter(pos, count) {
      auditMetrics.insertedCols.push({ pos, count });
      auditMetrics.maxCols += count;
    },
    deleteRows(pos, count) {
      auditMetrics.deletedRows.push({ pos, count });
      auditMetrics.maxRows -= count;
    },
    deleteColumns(pos, count) {
      auditMetrics.deletedCols.push({ pos, count });
      auditMetrics.maxCols -= count;
    },
    setRowHeight(r, h) { auditMetrics.rowHeights[r] = h; },
    setFrozenRows(n) { auditMetrics.frozenRows = n; },
    setColumnWidth(c, w) { auditMetrics.colWidths[c] = w; },
    getFilter() {
      if (!auditMetrics.hasFilter) return null;
      return {
        remove() {
          auditMetrics.hasFilter = false;
          auditMetrics.filterRemovedCount++;
        }
      };
    }
  }
};

const driveFixtureFiles = {
  FOLDER_TEMP_123: {
    id: 'FOLDER_TEMP_123',
    name: '_Doc2PDF_Temp',
    mimeType: 'application/vnd.google-apps.folder',
    trashed: false
  },
  CONTROL_SHEET_ID: {
    id: 'CONTROL_SHEET_ID',
    name: 'Doc2PDF Control',
    mimeType: 'application/vnd.google-apps.spreadsheet',
    trashed: false,
    parents: ['SHEET_PARENT_DIR'],
    owners: [{ emailAddress: 'admin@centro.edu' }],
    permissions: [{ type: 'user', emailAddress: 'admin@centro.edu', role: 'owner' }]
  },
  TEMP_OTHER_USER_FILE: {
    id: 'TEMP_OTHER_USER_FILE',
    name: '__tmp_JOB-X',
    mimeType: 'application/vnd.google-apps.document',
    trashed: false,
    parents: ['FOLDER_TEMP_123'],
    owners: [{ emailAddress: 'admin@centro.edu' }],
    permissions: [{ type: 'domain', domain: 'centro.edu', role: 'reader' }]
  },
  ADMIN_PRIVATE_DOC_12345: {
    id: 'ADMIN_PRIVATE_DOC_12345',
    name: 'Nomina_Confidencial.docx',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    size: '4096',
    trashed: false,
    parents: ['ADMIN_PRIV_DIR'],
    owners: [{ emailAddress: 'admin@centro.edu' }],
    permissions: [{ type: 'user', emailAddress: 'admin@centro.edu', role: 'owner' }]
  },
  NATIVE_GDOC_1234567890: {
    id: 'NATIVE_GDOC_1234567890',
    name: 'Acta Claustro',
    mimeType: 'application/vnd.google-apps.document',
    trashed: false,
    parents: ['SHARED_DIR'],
    owners: [{ emailAddress: 'docente@centro.edu' }],
    permissions: [{ type: 'user', emailAddress: 'docente@centro.edu', role: 'owner' }]
  },
  DESKTOP_XLSX_1234567890: {
    id: 'DESKTOP_XLSX_1234567890',
    name: 'Notas_Evaluacion.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    size: '4096',
    trashed: false,
    parents: ['SHARED_DIR'],
    owners: [{ emailAddress: 'otro@centro.edu' }],
    permissions: [{ type: 'domain', domain: 'centro.edu', role: 'reader' }]
  }
};

const sandbox = {
  console,
  Date,
  Math,
  JSON,
  Error,
  RegExp,
  Array,
  Object,
  String,
  Number,
  Boolean,
  encodeURIComponent,
  parseInt,
  isNaN,
  Utilities: {
    DigestAlgorithm: { SHA_256: 'SHA_256' },
    Charset: { UTF_8: 'UTF_8' },
    computeDigest(alg, text) {
      const buf = crypto.createHash('sha256').update(String(text)).digest();
      return Array.from(buf).map((b) => (b > 127 ? b - 256 : b));
    },
    formatDate(date, tz, fmt) {
      const d = new Date(date);
      const y = d.getUTCFullYear();
      const m = String(d.getUTCMonth() + 1).padStart(2, '0');
      const day = String(d.getUTCDate()).padStart(2, '0');
      const hh = String(d.getUTCHours()).padStart(2, '0');
      const mm = String(d.getUTCMinutes()).padStart(2, '0');
      const ss = String(d.getUTCSeconds()).padStart(2, '0');
      if (fmt === 'yyyy-MM-dd') return `${y}-${m}-${day}`;
      if (fmt === 'yyyyMMdd-HHmmss') return `${y}${m}${day}-${hh}${mm}${ss}`;
      if (fmt === 'dd/MM HH:mm') return `${day}/${m} ${hh}:${mm}`;
      if (fmt === 'dd/MM/yyyy HH:mm') return `${day}/${m}/${y} ${hh}:${mm}`;
      return d.toISOString();
    },
    base64Decode(b64) {
      return Array.from(Buffer.from(b64, 'base64'));
    },
    base64Encode(bytes) {
      return Buffer.from(bytes.map((b) => (b < 0 ? b + 256 : b))).toString('base64');
    },
    newBlob(bytes, mime, name) {
      return {
        getBytes: () => bytes,
        getContentType: () => mime,
        getName: () => name,
        setContentType(t) { mime = t; return this; },
        setName(n) { name = n; return this; }
      };
    },
    getUuid() {
      uuidCounter++;
      return 'uuid-' + uuidCounter;
    },
    sleep() {}
  },
  CacheService: {
    getScriptCache: () => ({
      get: (k) => (memCache.has(k) ? memCache.get(k) : null),
      put: (k, v, ttl) => {
        if (typeof ttl === 'number' && ttl > 21600) {
          throw new Error('CacheService TTL superó el límite real de 21600 s: ' + ttl);
        }
        memCache.set(k, v);
      },
      remove: (k) => memCache.delete(k)
    })
  },
  PropertiesService: {
    getScriptProperties: () => ({
      getProperty: (k) => (memProps.has(k) ? memProps.get(k) : null),
      setProperty: (k, v) => memProps.set(k, v),
      deleteProperty: (k) => memProps.delete(k)
    })
  },
  HtmlService: {
    createTemplateFromFile(fileName) {
      const tpl = {
        evaluate() {
          const rendered = {
            _width: null,
            _height: null,
            _file: fileName,
            _data: Object.assign({}, tpl)
          };
          return {
            setWidth(w) { rendered._width = w; return this; },
            setHeight(h) { rendered._height = h; return this; },
            _meta: rendered
          };
        }
      };
      return tpl;
    },
    createHtmlOutputFromFile(fileName) {
      return {
        getContent: () => fs.readFileSync(path.join(srcDir, `${fileName}.html`), 'utf8')
      };
    }
  },
  SpreadsheetApp: {
    BorderStyle: { SOLID: 'SOLID' },
    flush() {},
    getUi: () => ({
      ButtonSet: { OK: 'OK', YES_NO: 'YES_NO' },
      Button: { OK: 'OK', YES: 'YES', NO: 'NO' },
      alert(title, body, buttons) {
        uiAlerts.push({ title: String(title || ''), body: String(body || ''), buttons });
        return 'OK';
      },
      showModalDialog(htmlOutput, title) {
        uiModals.push({
          title: String(title || ''),
          meta: htmlOutput && htmlOutput._meta ? htmlOutput._meta : null
        });
      }
    }),
    newRichTextValue: () => {
      let text = '';
      let linkUrl = undefined;
      return {
        setText(t) { text = t; return this; },
        setLinkUrl(u) { linkUrl = u; return this; },
        build() { return { getText: () => text, getLinkUrl: () => linkUrl }; }
      };
    },
    newDataValidation: () => {
      const rule = {
        type: null,
        args: null,
        allowInvalid: true,
        helpText: ''
      };
      return {
        requireCheckbox(a, b) { rule.type = 'checkbox'; rule.args = [a, b]; return this; },
        requireValueInList(list, showDropdown) { rule.type = 'list'; rule.args = [list.slice(), showDropdown]; return this; },
        requireNumberBetween(min, max) { rule.type = 'number_between'; rule.args = [min, max]; return this; },
        setAllowInvalid(allow) { rule.allowInvalid = allow; return this; },
        setHelpText(txt) { rule.helpText = txt; return this; },
        build() { return Object.assign({}, rule); }
      };
    },
    getActiveSpreadsheet: () => ({
      getId: () => 'CONTROL_SHEET_ID',
      getOwner: () => ({ getEmail: () => 'admin@centro.edu' }),
      getEditors: () => [{ getEmail: () => 'coadmin@centro.edu' }],
      getSheets: () => [sheetsMock.Ajustes, sheetsMock['Auditoría']],
      getSheetByName: (name) => sheetsMock[name] || null,
      insertSheet: (name) => sheetsMock[name],
      getSpreadsheetLocale: () => 'es_ES',
      getSpreadsheetTimeZone: () => 'Europe/Madrid',
      toast() {}
    })
  },
  LockService: {
    getScriptLock: () => ({
      waitLock() {},
      releaseLock() {}
    })
  },
  Session: {
    getActiveUser: () => ({ getEmail: () => activeEmail }),
    getEffectiveUser: () => ({ getEmail: () => 'admin@centro.edu' }),
    getScriptTimeZone: () => 'Europe/Madrid'
  },
  GroupsApp: {
    getGroupByEmail(groupEmail) {
      groupsLookupCount++;
      const key = String(groupEmail || '').trim().toLowerCase();
      const spec = mockGroups[key];
      if (!spec) {
        throw new Error('Cannot find group: ' + key);
      }
      if (spec.throwLookup) {
        throw new Error(spec.throwLookup);
      }
      return {
        hasUser(userEmail) {
          groupsHasUserCount++;
          if (spec.throwHasUser) {
            throw new Error(spec.throwHasUser);
          }
          const clean = String(userEmail || '').trim().toLowerCase();
          return spec.members.has(clean);
        }
      };
    }
  },
  ScriptApp: {
    getOAuthToken: () => 'MOCK_TOKEN',
    getProjectTriggers: () => mockTriggers.slice(),
    deleteTrigger(t) {
      const idx = mockTriggers.indexOf(t);
      if (idx !== -1) mockTriggers.splice(idx, 1);
    },
    newTrigger: (fnName) => ({
      timeBased() { return this; },
      everyDays() { return this; },
      atHour() { return this; },
      create() {
        mockTriggers.push({ getHandlerFunction: () => String(fnName || 'runDailyMaintenance') });
      }
    })
  },
  Drive: {
    About: {
      get: () => mockAboutResponse
    },
    Files: {
      get(id) {
        if (driveFixtureFiles[id]) return driveFixtureFiles[id];
        throw new Error('404 File not found');
      },
      list() {
        return { files: mockListFiles.slice(), nextPageToken: null };
      },
      create(meta) {
        const id = 'TMP_GOOGLE_' + (createdDriveFiles.length + 1);
        createdDriveFiles.push({ id, meta });
        return { id };
      },
      copy(meta) {
        const id = 'TMP_COPY_' + (createdDriveFiles.length + 1);
        createdDriveFiles.push({ id, meta });
        return { id };
      },
      remove(id) {
        removedFiles.push(id);
        mockListFiles = mockListFiles.filter((f) => f.id !== id);
      }
    }
  },
  UrlFetchApp: {
    fetch(url, opts) {
      if (url.includes('uploadType=resumable')) {
        return {
          getResponseCode: () => 200,
          getAllHeaders: () => ({ Location: 'https://www.googleapis.com/upload/mock-session-1' }),
          getContentText: () => ''
        };
      }
      if (url === 'https://www.googleapis.com/upload/mock-session-1') {
        return {
          getResponseCode: () => 200,
          getAllHeaders: () => ({}),
          getContentText: () => JSON.stringify({ id: 'RESUMABLE_TMP_99' })
        };
      }
      if (url.includes('alt=media')) {
        return {
          getResponseCode: () => 206,
          getContent: () => Array.from(Buffer.from('%PDF-CHUNK')),
          getAllHeaders: () => ({})
        };
      }
      if (shouldFailExport) {
        return {
          getResponseCode: () => 500,
          getAllHeaders: () => ({ 'Content-Type': 'text/html' }),
          getBlob: () => null
        };
      }
      const pdfBytes = Array.from(Buffer.from('%PDF-1.4\n%MockPDF'));
      return {
        getResponseCode: () => 200,
        getAllHeaders: () => ({ 'Content-Type': 'application/pdf' }),
        getBlob: () => sandbox.Utilities.newBlob(pdfBytes, 'application/pdf', 'out.pdf')
      };
    }
  }
};

vm.createContext(sandbox);
for (const file of ['Config.gs', 'Cleanup.gs', 'Auth.gs', 'Audit.gs', 'Converter.gs', 'Setup.gs', 'WebApp.gs']) {
  const code = fs.readFileSync(path.join(srcDir, file), 'utf8');
  vm.runInContext(code, sandbox, { filename: file });
}

// 4.1 Extensiones, URLs, IDs, Sanitización
const expectedExts = [
  'docx', 'doc', 'dot', 'dotx', 'odt', 'rtf',
  'txt', 'md', 'html', 'htm',
  'xlsx', 'xls', 'xlt', 'xltx', 'xlsm', 'ods', 'csv', 'tsv',
  'pptx', 'ppt', 'odp'
];
const cfg = sandbox.getConfig_(true);
for (const ext of expectedExts) {
  assert(sandbox.getFormatByExtension_(ext, cfg), `Debe soportar .${ext}`);
}
assert.strictEqual(sandbox.getFormatByExtension_('jpg', cfg), null);
assert.strictEqual(sandbox.buildExportUrl_('DOC123', 'document'), 'https://docs.google.com/document/d/DOC123/export?format=pdf');
assert.strictEqual(sandbox.extractDriveFileId_('https://drive.google.com/file/d/1AbCdEfGhIjKlMnOpQrStUvWxYz/view'), '1AbCdEfGhIjKlMnOpQrStUvWxYz');
assert.strictEqual(sandbox.sanitizeCellValue_('=SUM(A1)'), '\'=SUM(A1)');

// 4.2 Flujo directo + borrado permanente incluso en error
const sampleB64 = Buffer.from('Dummy DOCX binary content').toString('base64');
const okRes = sandbox.convertLocalFileDirect_({ fileName: '=Notas.docx', base64Data: sampleB64 });
assert.strictEqual(okRes.transferMode, 'direct');
assert.strictEqual(removedFiles.length, 1);

shouldFailExport = true;
assert.throws(() => sandbox.convertLocalFileDirect_({ fileName: 'Fallo.docx', base64Data: sampleB64 }), /HTTP 500/);
shouldFailExport = false;
assert.strictEqual(removedFiles.length, 2);

// 4.3 Pruebas de intrusión: usuario normal llamando RPCs administrativos
activeEmail = 'docente@centro.edu';
for (const rpcCall of [
  () => sandbox.apiExecuteSetupSheet({ resetAudit: true }, 'fake'),
  () => sandbox.apiGetDashboardAnalytics(30, 'fake'),
  () => sandbox.apiPurgeTempFilesFromDialog('fake'),
  () => sandbox.apiInstallTriggerFromDialog('fake'),
  () => sandbox.apiRemoveTriggerFromDialog('fake'),
  () => sandbox.apiSaveWebAppUrl('https://script.google.com/macros/s/evil/exec', 'fake')
]) {
  const res = rpcCall();
  assert.strictEqual(res.success, false);
  assert.strictEqual(res.code, 'UNAUTHORIZED_ADMIN');
}
assert.throws(() => sandbox.runDailyMaintenance(), /no autorizada/i);
// Trigger real sí permitido
sandbox.runDailyMaintenance({ triggerUid: 'real-trigger-1' });

// 4.4 Admin con dialogToken ausente/inválido vs válido
activeEmail = 'admin@centro.edu';
const badTokRes = sandbox.apiGetDashboardAnalytics(30, '');
assert.strictEqual(badTokRes.success, false);
assert.strictEqual(badTokRes.code, 'INVALID_DIALOG_TOKEN');

const realDlgToken = sandbox.createAdminDialogToken_('admin@centro.edu');
const goodDashRes = sandbox.apiGetDashboardAnalytics(30, realDlgToken);
assert.strictEqual(goodDashRes.success, true);

// 4.5 Pruebas anti-Confused Deputy en apiConvertDriveUrl
activeEmail = 'docente@centro.edu';
const controlSheetAttempt = sandbox.apiConvertDriveUrl({ url: 'https://docs.google.com/spreadsheets/d/CONTROL_SHEET_ID/edit' });
assert.strictEqual(controlSheetAttempt.success, false);
assert.strictEqual(controlSheetAttempt.code, 'DRIVE_ACCESS_FORBIDDEN');

const tempFileAttempt = sandbox.apiConvertDriveUrl({ url: 'TEMP_OTHER_USER_FILE' });
assert.strictEqual(tempFileAttempt.success, false);
assert.strictEqual(tempFileAttempt.code, 'DRIVE_ACCESS_FORBIDDEN');

const privateAdminAttempt = sandbox.apiConvertDriveUrl({ url: 'ADMIN_PRIVATE_DOC_12345' });
assert.strictEqual(privateAdminAttempt.success, false);
assert.strictEqual(privateAdminAttempt.code, 'DRIVE_ACCESS_NOT_SHARED');

// Archivo compartido con dominio centro.edu -> permitido para docente
const domainSharedOk = sandbox.apiConvertDriveUrl({ url: 'DESKTOP_XLSX_1234567890' });
assert.strictEqual(domainSharedOk.success, true);
assert.strictEqual(domainSharedOk.data.pdfFileName, 'Notas_Evaluacion.pdf');

// Archivo privado del admin pedido por el propio admin -> permitido
activeEmail = 'admin@centro.edu';
const adminOwnOk = sandbox.apiConvertDriveUrl({ url: 'ADMIN_PRIVATE_DOC_12345' });
assert.strictEqual(adminOwnOk.success, true);

// 4.6 Fail-Closed sin email de sesión
activeEmail = '';
const noEmailBoot = sandbox.apiGetBootstrap();
assert.strictEqual(noEmailBoot.success, true);
assert.strictEqual(noEmailBoot.data.user.allowed, false);
assert.strictEqual(noEmailBoot.data.user.code, 'NO_SESSION_EMAIL');
const noEmailConv = sandbox.apiConvertLocalDirect({ fileName: 'a.docx', base64Data: sampleB64 });
assert.strictEqual(noEmailConv.success, false);
assert.strictEqual(noEmailConv.code, 'NO_SESSION_EMAIL');

// 4.7 Modo mantenimiento: docente bloqueado, admin bypass
serviceActiveSetting = 'NO';
memCache.clear();
activeEmail = 'docente@centro.edu';
const maintUser = sandbox.apiGetBootstrap();
assert.strictEqual(maintUser.data.user.allowed, false);
assert.strictEqual(maintUser.data.user.code, 'MAINTENANCE');

activeEmail = 'admin@centro.edu';
const maintAdmin = sandbox.apiGetBootstrap();
assert.strictEqual(maintAdmin.data.user.allowed, true);
assert.strictEqual(maintAdmin.data.user.code, 'MAINTENANCE_ADMIN_BYPASS');
serviceActiveSetting = 'SI';
memCache.clear();

// 4.8 Secuestro de sesión fragmentada entre usuarios
activeEmail = 'docente@centro.edu';
const rawChunk = Buffer.from('Chunked payload bytes');
const initUp = sandbox.apiInitChunkedUpload({ fileName: 'Memoria.pptx', totalBytes: rawChunk.length });
assert.strictEqual(initUp.success, true);

activeEmail = 'otro@centro.edu';
for (const hijackCall of [
  () => sandbox.apiUploadChunk({ uploadId: initUp.data.uploadId, start: 0, base64Chunk: rawChunk.toString('base64') }),
  () => sandbox.apiFinalizeChunkedUpload({ uploadId: initUp.data.uploadId }),
  () => sandbox.apiAbortChunkedUpload({ uploadId: initUp.data.uploadId })
]) {
  const res = hijackCall();
  assert.strictEqual(res.success, false);
  assert.strictEqual(res.code, 'SESSION_FORBIDDEN');
}

// Secuestro de token de descarga de PDF
memCache.set('DOC2PDF_DL_tok1', JSON.stringify({ fileId: 'PDF_TMP_1', totalBytes: 100, totalChunks: 1, pdfFileName: 'x.pdf', userEmail: 'docente@centro.edu' }));
for (const dlHijack of [
  () => sandbox.apiFetchPdfChunk({ downloadToken: 'tok1', chunkIndex: 0 }),
  () => sandbox.apiCompletePdfDownload({ downloadToken: 'tok1' })
]) {
  const res = dlHijack();
  assert.strictEqual(res.success, false);
  assert.strictEqual(res.code, 'SESSION_FORBIDDEN');
}

// Propietario legítimo sí completa su subida fragmentada
activeEmail = 'docente@centro.edu';
const upOk = sandbox.apiUploadChunk({ uploadId: initUp.data.uploadId, start: 0, base64Chunk: rawChunk.toString('base64') });
assert.strictEqual(upOk.success, true);
const finOk = sandbox.apiFinalizeChunkedUpload({ uploadId: initUp.data.uploadId });
assert.strictEqual(finOk.success, true);

// 4.9 Setup de Ajustes, validaciones, recorte y carpeta temporal (Opción B + salvaguarda purga)
activeEmail = 'admin@centro.edu';
memProps.set('TEMP_FOLDER_ID', 'FOLDER_TEMP_123');
ajustesMetrics.cellUpdates = { '22,2': '' };
memCache.clear();
const setupTok = sandbox.createAdminDialogToken_('admin@centro.edu');
const setupRes = sandbox.apiExecuteSetupSheet({ resetAudit: false }, setupTok);
assert.strictEqual(setupRes.success, true);

// Recorte de filas y columnas sobrantes en Ajustes
assert.strictEqual(sheetsMock.Ajustes.getMaxRows(), 22, 'Ajustes debe quedar con exactamente 22 filas');
assert.strictEqual(sheetsMock.Ajustes.getMaxColumns(), 3, 'Ajustes debe quedar con exactamente 3 columnas');
assert.strictEqual(ajustesMetrics.deletedRows[0].pos, 23);
assert.strictEqual(ajustesMetrics.deletedRows[0].count, 978);
assert.strictEqual(ajustesMetrics.deletedCols[0].pos, 4);
assert.strictEqual(ajustesMetrics.deletedCols[0].count, 23);

// Multilínea en columna B y autoanchura + holgura en columna C
assert(ajustesMetrics.clearDataValidationsCount >= 1);
assert(ajustesMetrics.wrapCalls.some((w) => w.col === 2 && w.numRows === 21 && w.wrap === true), 'Columna B debe tener wrap=true');
assert(ajustesMetrics.wrapCalls.some((w) => w.col === 3 && w.numRows === 21 && w.wrap === false), 'Columna C debe tener wrap=false antes de autoResizeColumn');
assert(ajustesMetrics.autoResizedCols.includes(3), 'Debe ejecutar autoResizeColumn(3)');
assert.strictEqual(ajustesMetrics.colWidths[3], 637, 'Columna C debe incluir holgura proporcional del +8% + 32 px');
assert(ajustesMetrics.rowHeights[2] >= 28, 'Altura mínima de fila >= 28');

// Validaciones en casillas de familias, VISOR_PDF_EMBEBIDO, desplegable auditoría y rangos numéricos
for (const chkRow of [14, 15, 16, 17, 19]) {
  assert.strictEqual(ajustesMetrics.validations[chkRow].type, 'checkbox', `Fila ${chkRow} debe tener casilla SI/NO`);
  assert.strictEqual(JSON.stringify(ajustesMetrics.validations[chkRow].args), JSON.stringify(['SI', 'NO']));
  assert.strictEqual(ajustesMetrics.validations[chkRow].allowInvalid, false);
}
assert.strictEqual(ajustesMetrics.validations[20].type, 'list', 'REGISTRAR_NOMBRE_ARCHIVO debe tener lista');
assert.strictEqual(
  JSON.stringify(ajustesMetrics.validations[20].args[0]),
  JSON.stringify(['COMPLETO', 'ANONIMIZADO', 'NINGUNO'])
);
assert.strictEqual(ajustesMetrics.validations[20].allowInvalid, false);

const numericAssertions = [
  [11, 0, 10000],  // LIMITE_DIARIO_USUARIO
  [12, 1, 45],     // TAMANO_MAX_MB
  [13, 1, 20],     // MAX_ARCHIVOS_LOTE
  [21, 0, 3650]    // RETENCION_AUDITORIA_DIAS
];
for (const [rowNum, minVal, maxVal] of numericAssertions) {
  assert.strictEqual(ajustesMetrics.validations[rowNum].type, 'number_between');
  assert.strictEqual(JSON.stringify(ajustesMetrics.validations[rowNum].args), JSON.stringify([minVal, maxVal]));
  assert.strictEqual(ajustesMetrics.validations[rowNum].allowInvalid, false);
}

// Normalización y filtrado con las 4 casillas independientes PERMITIR_* + VISOR_PDF_EMBEBIDO
const partialCfg = sandbox.normalizeConfigObject_({
  PERMITIR_DOCUMENTOS: 'SI',
  PERMITIR_HOJAS: 'NO',
  PERMITIR_PRESENTACIONES: 'SI',
  PERMITIR_TEXTO: 'NO',
  VISOR_PDF_EMBEBIDO: 'SI'
});
assert.strictEqual(
  JSON.stringify(partialCfg.FAMILIAS_HABILITADAS),
  JSON.stringify(['DOCUMENTOS', 'PRESENTACIONES'])
);
assert.strictEqual(partialCfg.VISOR_PDF_EMBEBIDO, true);
assert.strictEqual(sandbox.normalizeConfigObject_({}).VISOR_PDF_EMBEBIDO, false);
const pubCfg = sandbox.getPublicClientConfig_({ allowed: true }, partialCfg);
assert.strictEqual(pubCfg.embeddedPdfViewer, true);
assert.strictEqual(sandbox.getFormatByExtension_('xlsx', partialCfg), null);
assert.strictEqual(sandbox.getFormatByExtension_('docx', partialCfg).ext, 'docx');

// Opción B: _Doc2PDF_Temp creada dentro de la carpeta padre de la hoja ('SHEET_PARENT_DIR') y escrita en CARPETA_TEMP_ID
const createdTempFolder = createdDriveFiles.find((f) => f.meta && f.meta.mimeType === 'application/vnd.google-apps.folder');
assert(createdTempFolder, 'Debe haber creado _Doc2PDF_Temp');
assert.strictEqual(JSON.stringify(createdTempFolder.meta.parents), JSON.stringify(['SHEET_PARENT_DIR']));
assert.strictEqual(ajustesMetrics.cellUpdates['22,2'], createdTempFolder.id, 'CARPETA_TEMP_ID debe quedar rellenado en Ajustes');

// Sincronización de CARPETA_TEMP_ID cuando ya existe en ScriptProperties pero la celda está vacía
memProps.set('TEMP_FOLDER_ID', 'FOLDER_TEMP_123');
ajustesMetrics.cellUpdates['22,2'] = '';
const resolvedId = sandbox.getOrCreateTempFolder_({ CARPETA_TEMP_ID: '' });
assert.strictEqual(resolvedId, 'FOLDER_TEMP_123');
assert.strictEqual(ajustesMetrics.cellUpdates['22,2'], 'FOLDER_TEMP_123', 'Debe sincronizar CARPETA_TEMP_ID desde ScriptProperties si celda vacía');

// Salvaguarda anti-borrado en purgeOrphanTempFiles_
const removedBeforePurge = removedFiles.length;
mockListFiles = [
  { id: 'CONTROL_SHEET_ID', name: 'Doc2PDF Control', mimeType: 'application/vnd.google-apps.spreadsheet' },
  { id: 'SUBFOLDER_1', name: 'Carpeta', mimeType: 'application/vnd.google-apps.folder' },
  { id: 'ORPHAN_FILE_1', name: 'temp.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }
];
const purgeStats = sandbox.purgeOrphanTempFiles_(0);
assert.strictEqual(purgeStats.scanned, 1);
assert.strictEqual(purgeStats.deleted, 1);
assert.deepStrictEqual(removedFiles.slice(removedBeforePurge), ['ORPHAN_FILE_1']);
mockListFiles = [];

// 4.10 Recorte, cabeceras multilínea y filtros en Auditoría (hoja vacía tras reset vs. hoja existente con filas)
// Escenario A: Auditoría con filas existentes -> conserva filas, recorta columnas a 13, repara anchos y filtro
const rowsBeforePreserve = auditRows.length;
assert(rowsBeforePreserve >= 2, 'Debe haber registros previos en Auditoría');
auditMetrics.maxRows = 1000;
auditMetrics.maxCols = 26;
auditMetrics.hasFilter = true;
auditMetrics.deletedRows = [];
auditMetrics.deletedCols = [];
auditMetrics.wrapCalls = [];
const preserveRes = sandbox.apiExecuteSetupSheet({ resetAudit: false }, setupTok);
assert.strictEqual(preserveRes.success, true);
assert.strictEqual(sheetsMock['Auditoría'].getMaxColumns(), 13, 'Auditoría debe recortarse a 13 columnas');
assert.strictEqual(sheetsMock['Auditoría'].getMaxRows(), 1 + rowsBeforePreserve, 'Auditoría conserva sus filas registradas');
assert.strictEqual(auditMetrics.filterRemovedCount, 1, 'Debe eliminar filtro previo antes de recortar');
assert.strictEqual(auditMetrics.frozenRows, 1);
assert(auditMetrics.wrapCalls.some((w) => w.row === 1 && w.col === 1 && w.numCols === 13 && w.wrap === true), 'Fila 1 de Auditoría debe tener setWrap(true)');
assert.strictEqual(auditMetrics.colWidths[6], 118);
assert.strictEqual(auditMetrics.colWidths[8], 196);
assert.strictEqual(auditMetrics.colWidths[9], 168);
assert.strictEqual(auditMetrics.colWidths[10], 148);
assert.strictEqual(
  JSON.stringify(auditMetrics.filterRange),
  JSON.stringify({ row: 1, col: 1, numRows: 1 + rowsBeforePreserve, numCols: 13 })
);

// Escenario B: Auditoría con columnas extra de usuario (lastCol=15) y resetAudit=false -> solo borra columnas vacías 16..26
auditMetrics.maxRows = 1000;
auditMetrics.maxCols = 26;
auditMetrics.lastCol = 15;
auditMetrics.deletedRows = [];
auditMetrics.deletedCols = [];
const extraColRes = sandbox.apiExecuteSetupSheet({ resetAudit: false }, setupTok);
assert.strictEqual(extraColRes.success, true);
assert.strictEqual(sheetsMock['Auditoría'].getMaxColumns(), 15, 'Debe respetar getLastColumn() cuando usuario añadió datos hasta col 15');
assert.strictEqual(auditMetrics.deletedCols[0].pos, 16);
assert.strictEqual(auditMetrics.deletedCols[0].count, 11);

// Escenario C: resetAudit=true (Auditoría vacía) -> deja 2 filas (fila 1 + 1 fila libre bajo encabezado) y 13 columnas
auditMetrics.maxRows = 1000;
auditMetrics.maxCols = 26;
auditMetrics.deletedRows = [];
auditMetrics.deletedCols = [];
const resetRes = sandbox.apiExecuteSetupSheet({ resetAudit: true }, setupTok);
assert.strictEqual(resetRes.success, true);
assert.strictEqual(sheetsMock['Auditoría'].getMaxRows(), 2, 'Auditoría vacía debe conservar 2 filas para fijar fila 1');
assert.strictEqual(sheetsMock['Auditoría'].getMaxColumns(), 13, 'Auditoría vacía debe recortarse a 13 columnas');
assert.strictEqual(auditMetrics.deletedRows[0].pos, 3);
assert.strictEqual(auditMetrics.deletedRows[0].count, 998);
assert.strictEqual(auditMetrics.deletedCols[0].pos, 14);
assert.strictEqual(auditMetrics.deletedCols[0].count, 13);
assert.strictEqual(auditMetrics.frozenRows, 1);
assert.strictEqual(
  JSON.stringify(auditMetrics.filterRange),
  JSON.stringify({ row: 1, col: 1, numRows: 2, numCols: 13 })
);

// Verificar contenido de aviso explícito y spinner accesible en InitDialog.html + altura en Setup.gs
const initDlgHtml = fs.readFileSync(path.join(srcDir, 'InitDialog.html'), 'utf8');
assert(initDlgHtml.includes('cualquier dato añadido en filas o columnas adicionales'), 'InitDialog.html debe advertir sobre borrado de datos adicionales');
assert(initDlgHtml.includes('<progress id="initSpinner" class="loading-spinner"'), 'InitDialog.html debe usar <progress> nativo como spinner');
assert(initDlgHtml.includes('prefers-reduced-motion: reduce'), 'InitDialog.html debe respetar prefers-reduced-motion');
assert(initDlgHtml.includes('forced-colors: active'), 'InitDialog.html debe soportar forced-colors');
assert(initDlgHtml.includes('@supports not (background-clip: border-area)'), 'InitDialog.html debe incluir fallback mask-image para border-area');
const setupInitSrc = fs.readFileSync(path.join(srcDir, 'Setup.gs'), 'utf8');
assert(setupInitSrc.includes('.setWidth(510).setHeight(520)'), 'setupSheet debe usar altura 520px en InitDialog para evitar barra de scroll');

// 4.11 Extraer y probar generador ZIP nativo de ClientScript.html (CRC32 + STORE + UTF-8 + deduplicación)
const clientHtml = fs.readFileSync(path.join(srcDir, 'ClientScript.html'), 'utf8');
const zipStart = clientHtml.indexOf('const CRC32_TABLE =');
const zipEnd = clientHtml.indexOf('function buildZipBlob(entries)');
assert(zipStart !== -1 && zipEnd > zipStart, 'Debe existir el empaquetador ZIP nativo en ClientScript.html');
const zipCode = clientHtml.slice(zipStart, zipEnd);
const zipSandbox = { Uint32Array, Uint8Array, DataView, TextEncoder, Date, Math, Object, String, Number, unescape, encodeURIComponent };
vm.createContext(zipSandbox);
vm.runInContext(zipCode, zipSandbox, { filename: 'ClientScript_ZIP' });

const usedZipNames = Object.create(null);
const pdfA = Buffer.from('%PDF-1.4\nContenido UTF-8 con ñ y tildes\n%%EOF', 'utf8');
const pdfB = Buffer.from('%PDF-1.4\nPrimer archivo Notas\n%%EOF', 'utf8');
const pdfC = Buffer.from('%PDF-1.4\nSegundo archivo Notas\n%%EOF', 'utf8');
const zipEntries = [
  { name: zipSandbox.makeUniqueZipName('Evaluación_Año.pdf', usedZipNames), bytes: new Uint8Array(pdfA) },
  { name: zipSandbox.makeUniqueZipName('Notas.pdf', usedZipNames), bytes: new Uint8Array(pdfB) },
  { name: zipSandbox.makeUniqueZipName('Notas.pdf', usedZipNames), bytes: new Uint8Array(pdfC) }
];
assert.strictEqual(zipEntries[2].name, 'Notas (2).pdf');
const rawZip = zipSandbox.buildZipBytes(zipEntries);
fs.writeFileSync('/tmp/test_doc2pdf_batch.zip', Buffer.from(rawZip));

// 4.12 Cabeceras en español, formato de fecha regional y orden del botón Reconvertir
const headersInVm = vm.runInContext('AUDIT_HEADERS', sandbox);
assert.strictEqual(headersInVm[0], 'Marca temporal');
assert.strictEqual(headersInVm[1], 'ID trabajo');
assert.strictEqual(auditMetrics.colWidths[1], 170);
assert(auditMetrics.numberFormats.some((nf) => nf.col === 1 && nf.fmt === 'dd/mm/yyyy hh:mm:ss'), 'Columna 1 debe formatearse en dd/mm/yyyy hh:mm:ss para es_ES');
assert.strictEqual(sandbox.getRegionalDateTimeFormat_({ getSpreadsheetLocale: () => 'es_ES' }), 'dd/mm/yyyy hh:mm:ss');
assert.strictEqual(sandbox.getRegionalDateTimeFormat_({ getSpreadsheetLocale: () => 'en_US' }), 'mm/dd/yyyy hh:mm:ss');
assert.strictEqual(sandbox.getRegionalDateTimeFormat_({ getSpreadsheetLocale: () => 'ja_JP' }), 'yyyy/mm/dd hh:mm:ss');

const doneBlockStart = clientHtml.indexOf("if (job.status === 'done' && job.pdfUrl)");
const doneBlockEnd = clientHtml.indexOf("else if (job.status === 'error'", doneBlockStart);
const doneBlock = clientHtml.slice(doneBlockStart, doneBlockEnd);
assert(doneBlock.indexOf('Reconvertir') !== -1 && doneBlock.indexOf('Descargar') !== -1);
assert(doneBlock.indexOf('Reconvertir') < doneBlock.indexOf('Descargar'), 'Reconvertir debe ir antes que Descargar');

// 4.13 Texto plano sin hipervínculos en columna Archivo (p. ej. README.md), formato numérico regional y ajustes del Dashboard
auditMetrics.numberFormats = [];
auditMetrics.richTextCalls = [];
assert.strictEqual(sandbox.logAuditJob_({
  jobId: 'JOB-MD-1',
  userEmail: 'docente@centro.edu',
  origin: 'Local',
  fileName: 'README.md',
  ext: 'md',
  family: 'DOCUMENTOS',
  inputBytes: 1200,
  pdfBytes: 24000,
  durationMs: 3900,
  status: 'OK'
}), true);
assert(auditMetrics.numberFormats.some((nf) => nf.col === 5 && nf.fmt === '@'), 'logAuditJob_ debe formatear la columna Archivo como texto plano @');
assert.strictEqual(sandbox.defangAutoLinkText_('README.md'), 'README.\u200Bmd');
assert.strictEqual(sandbox.defangAutoLinkText_('README.\u200Bmd'), 'README.\u200Bmd', 'defangAutoLinkText_ debe ser idempotente');
assert.strictEqual(sandbox.defangAutoLinkText_('guia.es.docx'), 'guia.\u200Bes.\u200Bdocx');
assert.strictEqual(sandbox.defangAutoLinkText_('v1.2'), 'v1.2');
assert(
  auditMetrics.richTextCalls.some((rc) => rc.col === 5 && rc.values[0][0].getText() === 'README.\u200Bmd' && rc.values[0][0].getLinkUrl() === null),
  'logAuditJob_ debe neutralizar y anular hipervínculos automáticos en README.md'
);
const dashAnalyticsMd = sandbox.computeDashboardAnalytics_(auditRows, 30, new Date(), {});
assert(dashAnalyticsMd.recentJobs.some((j) => j.file === 'README.md'), 'Dashboard debe mostrar README.md limpio sin \\u200B');

auditMetrics.numberFormats = [];
auditMetrics.richTextCalls = [];
const repairMdRes = sandbox.apiExecuteSetupSheet({ resetAudit: false }, setupTok);
assert.strictEqual(repairMdRes.success, true);
assert(auditMetrics.numberFormats.some((nf) => nf.col === 5 && nf.row === 2 && nf.fmt === '@'), 'executeSetupSheet_ debe aplicar formato @ en columna Archivo existente');
assert(
  auditMetrics.richTextCalls.some((rc) => rc.col === 5 && rc.row === 2 && rc.values[0][0].getText() === 'README.\u200Bmd' && rc.values[0][0].getLinkUrl() === null),
  'executeSetupSheet_ debe neutralizar y limpiar enlaces existentes en columna Archivo de forma idempotente'
);

assert.strictEqual(sandbox.getSpreadsheetLocaleBcp47_({ getSpreadsheetLocale: () => 'es_ES' }), 'es-ES');
assert.strictEqual(sandbox.getSpreadsheetLocaleBcp47_({ getSpreadsheetLocale: () => 'en_US' }), 'en-US');
assert.strictEqual(pubCfg.locale, 'es-ES');

const dashboardHtml = fs.readFileSync(path.join(srcDir, 'DashboardDialog.html'), 'utf8');
assert(dashboardHtml.includes('google.script.host.close()'), 'DashboardDialog.html debe incluir botón Cerrar en la cabecera');
assert(dashboardHtml.includes('Intl.NumberFormat(DASH_LOCALE'), 'DashboardDialog.html debe formatear números con Intl.NumberFormat');
assert(dashboardHtml.includes('Percentil 95'), 'DashboardDialog.html debe explicar p95 mediante tooltip');
assert(clientHtml.includes('formatLocaleNumber') && clientHtml.includes('Intl.NumberFormat(loc'), 'ClientScript.html debe formatear bytes y segundos según la configuración regional');

const setupSrc = fs.readFileSync(path.join(srcDir, 'Setup.gs'), 'utf8');
assert(setupSrc.includes('.setWidth(1040).setHeight(960)'), 'showDashboardDialog debe solicitar dimensiones 1040x960');
assert(setupSrc.includes('template.locale = getSpreadsheetLocaleBcp47_'), 'showDashboardDialog debe pasar template.locale');

// 4.14 Visor PDF: ocultación de flechas en 1 página, estados desactivados en extremos, ajuste 100% ancho y página completa
const indexHtmlSrc = fs.readFileSync(path.join(srcDir, 'Index.html'), 'utf8');
assert(indexHtmlSrc.includes('class="btn-sm hidden" id="pdfPrevBtn"'), 'Index.html debe ocultar pdfPrevBtn por defecto');
assert(indexHtmlSrc.includes('class="btn-sm hidden" id="pdfNextBtn"'), 'Index.html debe ocultar pdfNextBtn por defecto');
assert(indexHtmlSrc.includes('id="pdfPrevBtn" aria-label="Página anterior" title="Página anterior" disabled'), 'pdfPrevBtn debe tener title y disabled');
assert(indexHtmlSrc.includes('id="pdfNextBtn" aria-label="Página siguiente" title="Página siguiente" disabled'), 'pdfNextBtn debe tener title y disabled');
assert(indexHtmlSrc.includes('id="pdfZoomFitPageBtn"') && indexHtmlSrc.includes('aria-pressed="false">Página completa</button>'), 'Index.html debe incluir botón Página completa con aria-pressed');

const stylesHtmlSrc = fs.readFileSync(path.join(srcDir, 'Styles.html'), 'utf8');
assert(stylesHtmlSrc.includes('.btn-sm:disabled'), 'Styles.html debe definir estilo .btn-sm:disabled');
assert(stylesHtmlSrc.includes('.btn-sm[aria-pressed="true"]:not(:disabled)'), 'Styles.html debe definir estilo de estado activo desde aria-pressed');
assert(stylesHtmlSrc.includes('scrollbar-gutter: stable;'), 'Styles.html debe reservar scrollbar-gutter: stable en .preview-stage');
assert(stylesHtmlSrc.includes('@media (hover: hover)') && stylesHtmlSrc.includes(':not(:disabled)'), 'Styles.html debe restringir hover a dispositivos hover y botones activos');
assert(stylesHtmlSrc.includes('@media (forced-colors: active)'), 'Styles.html debe contemplar modo forced-colors para botones desactivados');

const previewCtrlStart = clientHtml.indexOf('function updatePreviewControls()');
const previewCtrlEnd = clientHtml.indexOf('function destroyPreviewState()', previewCtrlStart);
const previewScaleStart = clientHtml.indexOf('function computePreviewScale(');
const previewScaleEnd = clientHtml.indexOf('function renderCurrentPdfPage()', previewScaleStart);
assert(previewCtrlStart !== -1 && previewCtrlEnd > previewCtrlStart, 'Debe existir updatePreviewControls en ClientScript.html');
assert(previewScaleStart !== -1 && previewScaleEnd > previewScaleStart, 'Debe existir computePreviewScale en ClientScript.html');

const makeMockBtn = () => {
  const classes = new Set(['hidden']);
  const attrs = {};
  return {
    disabled: true,
    attrs,
    setAttribute(k, v) { attrs[k] = String(v); },
    classList: {
      toggle(cls, force) {
        if (force) classes.add(cls);
        else classes.delete(cls);
      },
      contains(cls) { return classes.has(cls); }
    }
  };
};
const previewMockEls = {
  pdfPrevBtn: makeMockBtn(),
  pdfNextBtn: makeMockBtn(),
  pdfZoomOutBtn: makeMockBtn(),
  pdfZoomFitBtn: makeMockBtn(),
  pdfZoomFitPageBtn: makeMockBtn(),
  pdfZoomInBtn: makeMockBtn(),
  pdfPageInfo: { textContent: '' }
};
const previewSandbox = {
  Boolean,
  Math,
  CSS: { supports: () => true },
  window: { getComputedStyle: () => ({ paddingLeft: '14px', paddingRight: '14px', paddingTop: '14px', paddingBottom: '14px' }) },
  parseFloat,
  state: { preview: { doc: null, pageNum: 1, totalPages: 1, fitMode: 'width' } },
  document: { getElementById: (id) => previewMockEls[id] }
};
vm.createContext(previewSandbox);
vm.runInContext(
  clientHtml.slice(previewCtrlStart, previewCtrlEnd) + '\n' + clientHtml.slice(previewScaleStart, previewScaleEnd),
  previewSandbox,
  { filename: 'ClientScript_Preview' }
);

// 1 página -> flechas ocultas y desactivadas, aria-pressed según fitMode
previewSandbox.state.preview = { doc: {}, pageNum: 1, totalPages: 1, fitMode: 'width' };
previewSandbox.updatePreviewControls();
assert.strictEqual(previewMockEls.pdfPageInfo.textContent, 'Página 1 / 1');
assert.strictEqual(previewMockEls.pdfPrevBtn.classList.contains('hidden'), true);
assert.strictEqual(previewMockEls.pdfNextBtn.classList.contains('hidden'), true);
assert.strictEqual(previewMockEls.pdfPrevBtn.disabled, true);
assert.strictEqual(previewMockEls.pdfNextBtn.disabled, true);
assert.strictEqual(previewMockEls.pdfZoomFitBtn.attrs['aria-pressed'], 'true');
assert.strictEqual(previewMockEls.pdfZoomFitPageBtn.attrs['aria-pressed'], 'false');

// Varias páginas, página 1 -> ambas visibles, izquierda desactivada, derecha activa
previewSandbox.state.preview = { doc: {}, pageNum: 1, totalPages: 4, fitMode: 'page' };
previewSandbox.updatePreviewControls();
assert.strictEqual(previewMockEls.pdfPageInfo.textContent, 'Página 1 / 4');
assert.strictEqual(previewMockEls.pdfPrevBtn.classList.contains('hidden'), false);
assert.strictEqual(previewMockEls.pdfNextBtn.classList.contains('hidden'), false);
assert.strictEqual(previewMockEls.pdfPrevBtn.disabled, true);
assert.strictEqual(previewMockEls.pdfNextBtn.disabled, false);
assert.strictEqual(previewMockEls.pdfZoomFitBtn.attrs['aria-pressed'], 'false');
assert.strictEqual(previewMockEls.pdfZoomFitPageBtn.attrs['aria-pressed'], 'true');

// Varias páginas, página intermedia en modo manual -> ambos botones fit desmarcados
previewSandbox.state.preview = { doc: {}, pageNum: 2, totalPages: 4, fitMode: 'manual' };
previewSandbox.updatePreviewControls();
assert.strictEqual(previewMockEls.pdfPrevBtn.disabled, false);
assert.strictEqual(previewMockEls.pdfNextBtn.disabled, false);
assert.strictEqual(previewMockEls.pdfZoomFitBtn.attrs['aria-pressed'], 'false');
assert.strictEqual(previewMockEls.pdfZoomFitPageBtn.attrs['aria-pressed'], 'false');

// Varias páginas, última página -> ambas visibles, izquierda activa, derecha desactivada
previewSandbox.state.preview = { doc: {}, pageNum: 4, totalPages: 4, fitMode: 'width' };
previewSandbox.updatePreviewControls();
assert.strictEqual(previewMockEls.pdfPageInfo.textContent, 'Página 4 / 4');
assert.strictEqual(previewMockEls.pdfPrevBtn.classList.contains('hidden'), false);
assert.strictEqual(previewMockEls.pdfNextBtn.classList.contains('hidden'), false);
assert.strictEqual(previewMockEls.pdfPrevBtn.disabled, false);
assert.strictEqual(previewMockEls.pdfNextBtn.disabled, true);

// Verificación geométrica de computePreviewScale (stage 860x560 => innerW = 830, innerH = 530)
const mockStage = { clientWidth: 860, clientHeight: 560 };
const innerContentW = 860 - 28; // 832
const innerContentH = 560 - 28; // 532

const sVertWidth = previewSandbox.computePreviewScale({ width: 595, height: 842 }, mockStage, 'width', 1);
assert(Math.floor(595 * sVertWidth) <= innerContentW, 'Ajuste al ancho en A4 vertical no debe desbordar horizontalmente');
assert(Math.floor(595 * sVertWidth) >= innerContentW - 4, 'Ajuste al ancho debe aprovechar prácticamente el 100 % de la anchura');

const sVertPage = previewSandbox.computePreviewScale({ width: 595, height: 842 }, mockStage, 'page', 1);
assert(Math.floor(595 * sVertPage) <= innerContentW && Math.floor(842 * sVertPage) <= innerContentH, 'Página completa vertical debe caber sin scroll');

const sHorizPage = previewSandbox.computePreviewScale({ width: 1400, height: 600 }, mockStage, 'page', 1);
assert(Math.floor(1400 * sHorizPage) <= innerContentW && Math.floor(600 * sHorizPage) <= innerContentH, 'Página completa horizontal debe caber sin scroll');

const sVeryWideWidth = previewSandbox.computePreviewScale({ width: 2000, height: 800 }, mockStage, 'width', 1);
assert(sVeryWideWidth < 0.5 && Math.floor(2000 * sVeryWideWidth) <= innerContentW, 'Hojas muy anchas deben poder escalarse por debajo de 0.5 sin desbordar');

// 4.15 Check compartido de maquetación PDF para hojas de cálculo en ambas pestañas
const clientHtmlLatest = fs.readFileSync(path.join(srcDir, 'ClientScript.html'), 'utf8');
assert(!indexHtmlSrc.includes('driveIsSheetChk') && !stylesHtmlSrc.includes('drive-sheet-toggle') && !clientHtmlLatest.includes('driveIsSheetChk'), 'No deben quedar referencias huérfanas a driveIsSheetChk');
assert(stylesHtmlSrc.includes('.sheet-layout-toggle'), 'Styles.html debe incluir .sheet-layout-toggle');

const panelDrivePos = indexHtmlSrc.indexOf('id="panelDrive"');
const togglePos = indexHtmlSrc.indexOf('id="sheetLayoutToggle"');
const panelDriveClosePos = indexHtmlSrc.indexOf('</div>', indexHtmlSrc.indexOf('class="drive-hint"', panelDrivePos));
const sheetPanelPos = indexHtmlSrc.indexOf('id="sheetOptionsPanel"');
assert(panelDriveClosePos !== -1 && togglePos > panelDriveClosePos && togglePos < sheetPanelPos, 'sheetLayoutToggle debe ubicarse fuera de #panelDrive y antes de #sheetOptionsPanel');

const sheetVisStart = clientHtmlLatest.indexOf('function updateSheetOptionsVisibility()');
const sheetVisEnd = clientHtmlLatest.indexOf('/* -------------------------------------------------------------------------', sheetVisStart);
assert(sheetVisStart !== -1 && sheetVisEnd > sheetVisStart);
const sheetPanelClasses = new Set(['hidden']);
const sheetChkMock = { checked: false };
const sheetVisSandbox = {
  Array,
  state: { config: { enabledFamilies: ['DOCUMENTOS', 'HOJAS'] } },
  document: {
    getElementById: (id) => {
      if (id === 'sheetLayoutChk') return sheetChkMock;
      if (id === 'sheetOptionsPanel') {
        return {
          classList: {
            add: (c) => sheetPanelClasses.add(c),
            toggle: (c, f) => (f ? sheetPanelClasses.add(c) : sheetPanelClasses.delete(c))
          }
        };
      }
      return null;
    }
  }
};
vm.createContext(sheetVisSandbox);
vm.runInContext(clientHtmlLatest.slice(sheetVisStart, sheetVisEnd), sheetVisSandbox, { filename: 'ClientScript_SheetVis' });
sheetVisSandbox.updateSheetOptionsVisibility();
assert.strictEqual(sheetPanelClasses.has('hidden'), true, 'Con checkbox desmarcado el panel de maquetación debe estar oculto');
sheetChkMock.checked = true;
sheetVisSandbox.updateSheetOptionsVisibility();
assert.strictEqual(sheetPanelClasses.has('hidden'), false, 'Al marcar el checkbox el panel debe mostrarse en cualquier pestaña');
sheetVisSandbox.state.config.enabledFamilies = ['DOCUMENTOS'];
sheetVisSandbox.updateSheetOptionsVisibility();
assert.strictEqual(sheetPanelClasses.has('hidden'), true, 'Si HOJAS está deshabilitado en Ajustes, el panel debe permanecer oculto aunque se marque el check');

// 4.16 Corrección TTL cuota (21600 s) + zona horaria, exención admin e impersonación blindada
assert.strictEqual(vm.runInContext('DEV_IMPERSONATE_ENABLED_', sandbox), false, 'DEV_IMPERSONATE_ENABLED_ debe estar desactivado por defecto en producción y Git');
assert.strictEqual(vm.runInContext('DEV_IMPERSONATE_EMAIL_', sandbox), '', 'DEV_IMPERSONATE_EMAIL_ debe estar vacío por defecto en producción y Git');
assert.strictEqual(vm.runInContext('QUOTA_CACHE_TTL_SEC_', sandbox), 21600, 'TTL de caché de cuota debe respetar el límite real de 21600 s');

const authSrc = fs.readFileSync(path.join(srcDir, 'Auth.gs'), 'utf8');
assert(authSrc.includes('getLocalTimeZone_()') && authSrc.includes('getLocalTimeZone_(ss)'), 'Auth.gs debe usar la zona horaria local de la hoja');

const prevActiveEmail416 = activeEmail;
const savedAuditLen416 = auditRows.length;
memCache.clear();
const baselineDocenteUsage = sandbox.getUserDailyUsage_('docente@centro.edu');
memCache.clear();
auditRows.push(
  [new Date(), 'J-ADM-1', 'admin@centro.edu', 'Local', 'a.docx', 'docx', 'DOCUMENTOS', 10, 20, 400, 'OK', '', ''],
  [new Date(), 'J-ADM-2', 'admin@centro.edu', 'Local', 'b.docx', 'docx', 'DOCUMENTOS', 10, 20, 400, 'OK', '', ''],
  [new Date(), 'J-DOC-1', 'docente@centro.edu', 'Local', 'c.docx', 'docx', 'DOCUMENTOS', 10, 20, 400, 'OK', '', ''],
  [new Date(), 'J-SIM-1', 'docente@centro.edu [simulado por admin@centro.edu]', 'Local', 'd.docx', 'docx', 'DOCUMENTOS', 10, 20, 400, 'OK', '', '']
);

const quotaCfgStrict = sandbox.normalizeConfigObject_({ DOMINIOS_PERMITIDOS: 'centro.edu', LIMITE_DIARIO_USUARIO: 1, SERVICIO_ACTIVO: 'SI' });
assert.strictEqual(sandbox.getUserDailyUsage_('docente@centro.edu'), baselineDocenteUsage + 1, 'Solo las conversiones reales cuentan para cuota; las simuladas no contaminan Auditoría');

activeEmail = 'admin@centro.edu';
const adminQuotaCtx = sandbox.getUserContext_(quotaCfgStrict);
assert.strictEqual(adminQuotaCtx.allowed, true, 'Los administradores deben quedar exentos de LIMITE_DIARIO_USUARIO');
assert.strictEqual(adminQuotaCtx.isAdmin, true);
assert.strictEqual(adminQuotaCtx.quotaExempt, true);
assert.strictEqual(adminQuotaCtx.remainingToday, null);

activeEmail = 'docente@centro.edu';
const userBlockedQuota = sandbox.getUserContext_(quotaCfgStrict);
assert.strictEqual(userBlockedQuota.allowed, false);
assert.strictEqual(userBlockedQuota.code, 'QUOTA_EXCEEDED');

// Impersonación estrictamente bloqueada para usuarios normales y co-editores
const quotaCfgNormal = sandbox.normalizeConfigObject_({ DOMINIOS_PERMITIDOS: 'centro.edu', LIMITE_DIARIO_USUARIO: 5, SERVICIO_ACTIVO: 'SI' });
activeEmail = 'docente@centro.edu';
assert.strictEqual(sandbox.getUserContext_(quotaCfgNormal, { enabled: true, email: 'otro@centro.edu' }).isImpersonating, false, 'Un usuario normal no puede impersonar');
activeEmail = 'coadmin@centro.edu';
assert.strictEqual(sandbox.getUserContext_(quotaCfgNormal, { enabled: true, email: 'docente@centro.edu' }).isImpersonating, false, 'Un co-editor distinto del desplegador efectivo no puede impersonar');

// Impersonación activa únicamente cuando real === effective (admin@centro.edu)
activeEmail = 'admin@centro.edu';
const simCtx = sandbox.getUserContext_(quotaCfgNormal, { enabled: true, email: 'docente@centro.edu' });
assert.strictEqual(simCtx.isImpersonating, true);
assert.strictEqual(simCtx.allowed, true);
assert.strictEqual(simCtx.email, 'docente@centro.edu');
assert.strictEqual(simCtx.realEmail, 'admin@centro.edu');
assert.strictEqual(simCtx.isAdmin, false);
assert.strictEqual(simCtx.quotaExempt, true, 'Las conversiones bajo impersonación no consumen cuota');
assert.strictEqual(simCtx.auditEmail, 'docente@centro.edu [simulado por admin@centro.edu]');
assert(simCtx.impersonationBanner.includes('docente@centro.edu'), 'Debe generarse banner visible de impersonación');

const quotaCfgMaint = sandbox.normalizeConfigObject_({ DOMINIOS_PERMITIDOS: 'centro.edu', LIMITE_DIARIO_USUARIO: 5, SERVICIO_ACTIVO: 'NO' });
const simMaintCtx = sandbox.getUserContext_(quotaCfgMaint, { enabled: true, email: 'docente@centro.edu' });
assert.strictEqual(simMaintCtx.allowed, false);
assert.strictEqual(simMaintCtx.code, 'MAINTENANCE', 'Impersonar a un usuario estándar permite verificar bloqueo de mantenimiento');
assert.strictEqual(simMaintCtx.isImpersonating, true);

// Restaurar estado previo
auditRows.length = savedAuditLen416;
activeEmail = prevActiveEmail416;
memCache.clear();

const converterSrc = fs.readFileSync(path.join(srcDir, 'Converter.gs'), 'utf8');
assert(converterSrc.includes('if (!userCtx.quotaExempt)') && converterSrc.includes('if (!session.quotaExempt)'), 'Converter.gs no debe incrementar cuota cuando quotaExempt es true');
assert(converterSrc.includes('userCtx.auditEmail || userCtx.displayEmail'), 'Converter.gs debe registrar auditEmail en Auditoría');
assert(converterSrc.includes('userCtx.realEmail || userCtx.email'), 'Converter.gs debe vincular sesiones de subida/descarga al realEmail');

assert(indexHtmlSrc.includes('id="impersonationBanner"'), 'Index.html debe incluir el banner visible de impersonación');
assert(stylesHtmlSrc.includes('.impersonation-banner'), 'Styles.html debe estilar .impersonation-banner');
assert(stylesHtmlSrc.includes('max-width: min(460px, 100%);'), 'Styles.html debe ampliar .user-pill para no truncar el correo innecesariamente');
assert(clientHtmlLatest.includes('Sin límite (Admin)') && clientHtmlLatest.includes('state.config.user.quotaExempt'), 'ClientScript.html debe mostrar exención admin y no decrementar cuota si quotaExempt');
assert(clientHtmlLatest.includes('emailEl.title ='), 'ClientScript.html debe asignar el correo completo en title');

const readmeLatest = fs.readFileSync(path.join(srcDir, 'README.md'), 'utf8');
assert(readmeLatest.includes('> [!NOTE]') && readmeLatest.includes('Sin límite (Admin)') && readmeLatest.includes('DEV_IMPERSONATE_ENABLED_'), 'README.md debe documentar alerta nativa, exención admin e impersonación');
assert(
  readmeLatest.includes('### Control de acceso: listas de usuarios y grupos') &&
  readmeLatest.includes('`g:grupo@dominio`') &&
  readmeLatest.includes('Gestionar listas de acceso y grupos'),
  'README.md debe documentar listas de usuarios y grupos, prefijo g: y herramienta de gestión'
);

// 4.17 Control de acceso con Google Groups, prefijo g:, precedencia lista negra > blanca, caché y diagnóstico
console.log('4.17 Verificando control de acceso mixto (usuarios + Google Groups)...');
const savedAuditLen417 = auditRows.length;
const prevActiveEmail417 = activeEmail;
const baseAccessCfg = {
  ...sandbox.getConfig_(true),
  SERVICIO_ACTIVO: true,
  DOMINIOS_PERMITIDOS: ['centro.edu'],
  LIMITE_DIARIO_USUARIO: 30
};
const resetGroupMocks = () => {
  for (const k of Object.keys(mockGroups)) delete mockGroups[k];
  mockGroups['claustro@centro.edu'] = { members: new Set(['docente@centro.edu', 'prof2@centro.edu']) };
  mockGroups['sancionados@centro.edu'] = { members: new Set(['sancionado@centro.edu', 'admin@centro.edu']) };
  mockGroups['restringido@centro.edu'] = { members: new Set(['alguien@centro.edu']), throwHasUser: 'You do not have permission to view members' };
  groupsLookupCount = 0;
  groupsHasUserCount = 0;
  memCache.clear();
};
resetGroupMocks();

// Orden de filas en DEFAULT_CONFIG_DEFS: LISTA_NEGRA_USUARIOS antes que LISTA_BLANCA_USUARIOS
const cfgDefKeys = Array.from(vm.runInContext('DEFAULT_CONFIG_DEFS.map(function (r) { return r[0]; })', sandbox));
assert(
  cfgDefKeys.indexOf('LISTA_NEGRA_USUARIOS') < cfgDefKeys.indexOf('LISTA_BLANCA_USUARIOS'),
  'LISTA_NEGRA_USUARIOS debe figurar antes que LISTA_BLANCA_USUARIOS en Ajustes'
);

// Normalización de listas con prefijos g: y grupo:
const parsedAccess = sandbox.parseAccessListSetting_(
  ' g: claustro@centro.edu, docente@centro.edu; grupo: equipo@centro.edu\nG: claustro@centro.edu '
);
assert.strictEqual(
  JSON.stringify(Array.from(parsedAccess)),
  JSON.stringify(['g:claustro@centro.edu', 'docente@centro.edu', 'g:equipo@centro.edu']),
  'parseAccessListSetting_ debe normalizar g:/grupo: con espacios y deduplicar'
);

// Coincidencia directa exacta => 0 llamadas a GroupsApp
activeEmail = 'docente@centro.edu';
let ctxGrp = sandbox.getUserContext_({
  ...baseAccessCfg,
  LISTA_NEGRA_USUARIOS: [],
  LISTA_BLANCA_USUARIOS: ['docente@centro.edu', 'g:claustro@centro.edu']
});
assert.strictEqual(ctxGrp.allowed, true);
assert.strictEqual(groupsLookupCount, 0, 'Coincidencia directa debe evitar consultar GroupsApp');

// Autodetección sin prefijo: cuenta usuario se cachea como NOT_GROUP y grupo válido autoriza y cachea membresía
resetGroupMocks();
activeEmail = 'docente@centro.edu';
const mixedCfg = {
  ...baseAccessCfg,
  LISTA_NEGRA_USUARIOS: [],
  LISTA_BLANCA_USUARIOS: ['usuario.suelto@centro.edu', 'claustro@centro.edu']
};
ctxGrp = sandbox.getUserContext_(mixedCfg);
assert.strictEqual(ctxGrp.allowed, true);
assert.strictEqual(groupsLookupCount, 2, 'Primera evaluación consulta ambos candidatos hasta confirmar miembro');
assert.strictEqual(groupsHasUserCount, 1);

// Segunda petición del mismo usuario usa caché GMEM (0 consultas nuevas)
ctxGrp = sandbox.getUserContext_(mixedCfg);
assert.strictEqual(ctxGrp.allowed, true);
assert.strictEqual(groupsLookupCount, 2, 'Segunda petición debe reutilizar caché de membresía');

// Otro usuario no miembro: usuario.suelto@centro.edu ya está en caché como NOT_GROUP, solo consulta claustro@centro.edu
activeEmail = 'externo@centro.edu';
ctxGrp = sandbox.getUserContext_(mixedCfg);
assert.strictEqual(ctxGrp.allowed, false);
assert.strictEqual(ctxGrp.code, 'NOT_IN_WHITELIST');
assert.strictEqual(groupsLookupCount, 3, 'NOT_GROUP cacheado evita reconsultar usuario.suelto@centro.edu');
assert.strictEqual(groupsHasUserCount, 2);

// Precedencia LISTA_NEGRA_USUARIOS > LISTA_BLANCA_USUARIOS y > Admin
resetGroupMocks();
activeEmail = 'sancionado@centro.edu';
ctxGrp = sandbox.getUserContext_({
  ...baseAccessCfg,
  LISTA_NEGRA_USUARIOS: ['g:sancionados@centro.edu'],
  LISTA_BLANCA_USUARIOS: ['sancionado@centro.edu', 'g:claustro@centro.edu']
});
assert.strictEqual(ctxGrp.allowed, false);
assert.strictEqual(ctxGrp.code, 'USER_BLOCKED', 'Grupo en lista negra prevalece sobre usuario en lista blanca');

activeEmail = 'admin@centro.edu';
ctxGrp = sandbox.getUserContext_({
  ...baseAccessCfg,
  LISTA_NEGRA_USUARIOS: ['g:sancionados@centro.edu'],
  LISTA_BLANCA_USUARIOS: []
});
assert.strictEqual(ctxGrp.allowed, false);
assert.strictEqual(ctxGrp.code, 'USER_BLOCKED', 'Incluso administradores quedan bloqueados si pertenecen a LISTA_NEGRA_USUARIOS');

// Admin exento de LISTA_BLANCA_USUARIOS cuando no está en lista negra
ctxGrp = sandbox.getUserContext_({
  ...baseAccessCfg,
  LISTA_NEGRA_USUARIOS: [],
  LISTA_BLANCA_USUARIOS: ['g:solo-docentes@centro.edu']
});
assert.strictEqual(ctxGrp.allowed, true, 'Admin sigue exento de LISTA_BLANCA_USUARIOS');

// Política Fail-Closed en prefijo explícito g: inexistente o inaccesible
resetGroupMocks();
activeEmail = 'docente@centro.edu';
ctxGrp = sandbox.getUserContext_({
  ...baseAccessCfg,
  LISTA_NEGRA_USUARIOS: ['g:inexistente@centro.edu'],
  LISTA_BLANCA_USUARIOS: []
});
assert.strictEqual(ctxGrp.allowed, false);
assert.strictEqual(ctxGrp.code, 'BLACKLIST_GROUP_ERROR');
assert(ctxGrp.reason.includes('g:inexistente@centro.edu'));

ctxGrp = sandbox.getUserContext_({
  ...baseAccessCfg,
  LISTA_NEGRA_USUARIOS: [],
  LISTA_BLANCA_USUARIOS: ['g:inexistente@centro.edu']
});
assert.strictEqual(ctxGrp.allowed, false);
assert.strictEqual(ctxGrp.code, 'WHITELIST_GROUP_ERROR');

// Política Fail-Closed cuando el grupo existe pero la lista de miembros no es legible (hasUser lanza error)
resetGroupMocks();
activeEmail = 'docente@centro.edu';
ctxGrp = sandbox.getUserContext_({
  ...baseAccessCfg,
  LISTA_NEGRA_USUARIOS: ['restringido@centro.edu'],
  LISTA_BLANCA_USUARIOS: []
});
assert.strictEqual(ctxGrp.allowed, false);
assert.strictEqual(ctxGrp.code, 'BLACKLIST_GROUP_ERROR', 'Error de permisos en miembros de grupo debe bloquear Fail-Closed');

ctxGrp = sandbox.getUserContext_({
  ...baseAccessCfg,
  LISTA_NEGRA_USUARIOS: [],
  LISTA_BLANCA_USUARIOS: ['restringido@centro.edu']
});
assert.strictEqual(ctxGrp.allowed, false);
assert.strictEqual(ctxGrp.code, 'WHITELIST_GROUP_ERROR', 'Error de permisos en lista blanca debe devolver WHITELIST_GROUP_ERROR');

// Invalidación de caché de grupos mediante clearConfigCache_()
resetGroupMocks();
activeEmail = 'prof2@centro.edu';
const dynCfg = {
  ...baseAccessCfg,
  LISTA_NEGRA_USUARIOS: [],
  LISTA_BLANCA_USUARIOS: ['g:claustro@centro.edu']
};
assert.strictEqual(sandbox.getUserContext_(dynCfg).allowed, true);
mockGroups['claustro@centro.edu'].members.delete('prof2@centro.edu');
assert.strictEqual(sandbox.getUserContext_(dynCfg).allowed, true, 'Caché de 5 min mantiene estado previo hasta expiración o limpieza');
sandbox.clearConfigCache_();
const afterFlushCtx = sandbox.getUserContext_(dynCfg);
assert.strictEqual(afterFlushCtx.allowed, false);
assert.strictEqual(afterFlushCtx.code, 'NOT_IN_WHITELIST', 'clearConfigCache_() debe invalidar época de caché de grupos al instante');

// Diagnóstico inspectAccessLists_() + menuCheckAccessLists()
resetGroupMocks();
customBlacklistValue = 'solapado@centro.edu, g:inexistente@centro.edu';
customWhitelistValue = 'g:claustro@centro.edu, restringido@centro.edu, solapado@centro.edu';
activeEmail = 'docente@centro.edu';
assert.throws(
  () => sandbox.menuCheckAccessLists(),
  /Operación administrativa no autorizada/,
  'menuCheckAccessLists debe bloquear a usuarios no administradores'
);

activeEmail = 'admin@centro.edu';
uiModals.length = 0;
sandbox.menuCheckAccessLists();
assert.strictEqual(uiModals.length, 1, 'menuCheckAccessLists debe abrir un modal HTML para administradores');
assert(uiModals[0].meta && uiModals[0].meta._file === 'AccessCheckDialog', 'Debe usar la plantilla AccessCheckDialog');
assert.strictEqual(uiModals[0].meta._width, 660);
assert.strictEqual(uiModals[0].meta._height, 640);
const modalReport = uiModals[0].meta._data.report;
assert.strictEqual(modalReport.overlaps.length, 1);
assert.strictEqual(modalReport.overlaps[0], 'solapado@centro.edu');
assert.strictEqual(modalReport.errors, 2);

const diagReport = sandbox.inspectAccessLists_(sandbox.getConfig_(true));
assert.strictEqual(diagReport.overlaps.length, 1);
assert.strictEqual(diagReport.overlaps[0], 'solapado@centro.edu');
assert.strictEqual(diagReport.errors, 2);

const accessDlgHtml = fs.readFileSync(path.join(srcDir, 'AccessCheckDialog.html'), 'utf8');
for (const token of [
  '<title>', '<header', '<main', '<footer', 'role="list"',
  "include_('DialogStyles')", 'google.script.host.close()'
]) {
  assert(accessDlgHtml.includes(token), `AccessCheckDialog.html debe incluir ${token}`);
}
const dialogStylesHtml = fs.readFileSync(path.join(srcDir, 'DialogStyles.html'), 'utf8');
for (const token of [
  'linear-gradient(135deg, var(--accent) 0%, #0f2b5b 100%)',
  'color: #ffffff;',
  ':focus-visible',
  '<style>'
]) {
  assert(dialogStylesHtml.includes(token), `DialogStyles.html debe incluir ${token}`);
}
const aboutDlgHtml = fs.readFileSync(path.join(srcDir, 'AboutDialog.html'), 'utf8');
assert(aboutDlgHtml.includes('Google Groups y auditoría'), 'AboutDialog.html debe incluir píldora Google Groups y auditoría');

// 4.18 Modales de mantenimiento (TriggerDialog, FormatsCheckDialog, PurgeTempDialog) + gestión de trigger
console.log('4.18 Verificando modales HTML de mantenimiento y ciclo de vida de trigger...');
const prevActiveEmail418 = activeEmail;
mockTriggers.length = 0;

// Bloqueo a no administradores en los 3 menús
activeEmail = 'docente@centro.edu';
for (const menuFn of [
  () => sandbox.menuInstallMaintenanceTrigger(),
  () => sandbox.menuCheckDriveImportFormats(),
  () => sandbox.menuPurgeTempFiles()
]) {
  assert.throws(menuFn, /Operación administrativa no autorizada/);
}

// TriggerDialog: estado inicial inactivo -> instalar -> activo -> desactivar -> inactivo
activeEmail = 'admin@centro.edu';
const trigTok = sandbox.createAdminDialogToken_('admin@centro.edu');
uiModals.length = 0;
sandbox.menuInstallMaintenanceTrigger();
assert.strictEqual(uiModals.length, 1);
assert.strictEqual(uiModals[0].meta._file, 'TriggerDialog');
assert.strictEqual(uiModals[0].meta._width, 540);
assert.strictEqual(uiModals[0].meta._height, 500);
assert.strictEqual(uiModals[0].meta._data.status.installed, false);
assert.strictEqual(uiModals[0].meta._data.status.count, 0);

const instTrigRes = sandbox.apiInstallTriggerFromDialog(trigTok);
assert.strictEqual(instTrigRes.success, true);
assert.strictEqual(instTrigRes.status.installed, true);
assert.strictEqual(instTrigRes.status.count, 1);
assert.strictEqual(sandbox.hasMaintenanceTrigger_(), true);

uiModals.length = 0;
sandbox.menuInstallMaintenanceTrigger();
assert.strictEqual(uiModals[0].meta._data.status.installed, true);
assert.strictEqual(uiModals[0].meta._data.status.count, 1);

const remTrigRes = sandbox.apiRemoveTriggerFromDialog(trigTok);
assert.strictEqual(remTrigRes.success, true);
assert.strictEqual(remTrigRes.result.removed, 1);
assert.strictEqual(remTrigRes.status.installed, false);
assert.strictEqual(sandbox.hasMaintenanceTrigger_(), false);
assert.strictEqual(sandbox.apiRemoveTriggerFromDialog('').code, 'INVALID_DIALOG_TOKEN');

// FormatsCheckDialog: agrupación por familia + maxImportMb
uiModals.length = 0;
sandbox.menuCheckDriveImportFormats();
assert.strictEqual(uiModals.length, 1);
assert.strictEqual(uiModals[0].meta._file, 'FormatsCheckDialog');
assert.strictEqual(uiModals[0].meta._width, 620);
assert.strictEqual(uiModals[0].meta._height, 600);
const fmtReport = uiModals[0].meta._data.report;
assert.strictEqual(fmtReport.totalEnabled, 21);
assert.strictEqual(fmtReport.verifiedCount, 6); // docx, doc, dot, xlsx, pptx, txt en el mock
assert.strictEqual(fmtReport.warningCount, 15);
assert.strictEqual(fmtReport.families.length, 4);
const docItems = fmtReport.families.find((f) => f.key === 'DOCUMENTOS').items;
assert.strictEqual(docItems.find((i) => i.ext === 'docx').maxImportMb, 50);

// PurgeTempDialog: recuento inicial, purga selectiva > 30 min, purga total y compatibilidad 1 arg (DashboardDialog)
memProps.set('TEMP_FOLDER_ID', 'FOLDER_TEMP_123');
const oldIso = new Date(Date.now() - 120 * 60 * 1000).toISOString();
const freshIso = new Date(Date.now() - 2 * 60 * 1000).toISOString();
mockListFiles = [
  { id: 'OLD_1', name: 'viejo.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', modifiedTime: oldIso },
  { id: 'FRESH_1', name: 'reciente.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', modifiedTime: freshIso }
];

uiModals.length = 0;
sandbox.menuPurgeTempFiles();
assert.strictEqual(uiModals.length, 1);
assert.strictEqual(uiModals[0].meta._file, 'PurgeTempDialog');
assert.strictEqual(uiModals[0].meta._width, 540);
assert.strictEqual(uiModals[0].meta._height, 500);
assert.strictEqual(uiModals[0].meta._data.initialCount, 2);

const partialPurge = sandbox.apiPurgeTempFilesFromDialog(30, trigTok);
assert.strictEqual(partialPurge.success, true);
assert.strictEqual(partialPurge.result.scanned, 2);
assert.strictEqual(partialPurge.result.deleted, 1);
assert.strictEqual(partialPurge.remaining, 1);

const compatFullPurge = sandbox.apiPurgeTempFilesFromDialog(trigTok);
assert.strictEqual(compatFullPurge.success, true);
assert.strictEqual(compatFullPurge.maxAgeMinutes, 0);
assert.strictEqual(compatFullPurge.result.deleted, 1);
assert.strictEqual(compatFullPurge.remaining, 0);
mockListFiles = [];

// Verificaciones estáticas de accesibilidad y coherencia HTML en los nuevos diálogos
for (const dlgFile of ['TriggerDialog.html', 'FormatsCheckDialog.html', 'PurgeTempDialog.html']) {
  const html = fs.readFileSync(path.join(srcDir, dlgFile), 'utf8');
  for (const reqToken of ['<title>', '<header', '<main', '<footer', "include_('DialogStyles')", 'google.script.host.close()']) {
    assert(html.includes(reqToken), `${dlgFile} debe incluir ${reqToken}`);
  }
}
const triggerHtml = fs.readFileSync(path.join(srcDir, 'TriggerDialog.html'), 'utf8');
assert(triggerHtml.includes('role="list"') && triggerHtml.includes('<progress id="trigSpinner"'), 'TriggerDialog.html debe incluir role="list" y spinner <progress>');
assert(triggerHtml.includes('class="btn btn-close"') && !triggerHtml.includes('class="btn-outline"'), 'TriggerDialog.html debe usar class="btn btn-close"');

const formatsHtml = fs.readFileSync(path.join(srcDir, 'FormatsCheckDialog.html'), 'utf8');
assert(
  formatsHtml.includes('role="list"') &&
  formatsHtml.includes('class="btn btn-close"') &&
  formatsHtml.includes('<details') &&
  formatsHtml.includes('<summary'),
  'FormatsCheckDialog.html debe usar <details>/<summary>, role="list" y btn-close'
);
assert(
  !/<summary[^>]*>[\s\S]*?<h[1-6]/i.test(formatsHtml),
  'FormatsCheckDialog.html no debe colocar encabezados dentro de <summary>'
);

const purgeHtml = fs.readFileSync(path.join(srcDir, 'PurgeTempDialog.html'), 'utf8');
assert(purgeHtml.includes('<fieldset') && purgeHtml.includes('for="optOrphans"') && purgeHtml.includes('for="optAll"'), 'PurgeTempDialog.html debe usar fieldset + label for/id');
assert(purgeHtml.includes('class="btn btn-close"') && !purgeHtml.includes('class="btn-outline"'), 'PurgeTempDialog.html debe usar class="btn btn-close"');
assert(purgeHtml.indexOf('id="purgeBtn"') < purgeHtml.indexOf('id="closeBtn"'), 'PurgeTempDialog.html debe colocar Ejecutar purga a la izquierda y Cerrar a la derecha');

// Homogeneidad de botones Cerrar / Cancelar en todos los diálogos de Sheets y Dashboard
const dialogStylesAll = fs.readFileSync(path.join(srcDir, 'DialogStyles.html'), 'utf8');
assert(
  dialogStylesAll.includes('border-radius: 8px;') &&
  dialogStylesAll.includes('.btn-close, .btn-outline') &&
  dialogStylesAll.includes('border-color: #cbd5e1;') &&
  dialogStylesAll.includes('.body > *') &&
  dialogStylesAll.includes('flex-shrink: 0;') &&
  !/\.btn-close\s*\{[^}]*999px/.test(dialogStylesAll),
  'DialogStyles.html debe unificar .btn-close al estilo blanco rectangular (8px, #cbd5e1) y evitar encogimiento en .body'
);

for (const alignedFile of ['AboutDialog.html', 'DeployDialog.html', 'InitDialog.html', 'Setup.gs']) {
  const src = fs.readFileSync(path.join(srcDir, alignedFile), 'utf8');
  assert(
    src.includes('footer-bar') && src.includes('google.script.host.close()'),
    `${alignedFile} debe usar .footer-bar con google.script.host.close()`
  );
}
const dashLatestHtml = fs.readFileSync(path.join(srcDir, 'DashboardDialog.html'), 'utf8');
assert(
  dashLatestHtml.includes('onclick="google.script.host.close()"') && dashLatestHtml.includes('Cerrar panel'),
  'DashboardDialog.html debe incluir botón Cerrar en la cabecera fija'
);

activeEmail = prevActiveEmail418;

// 4.19 Explorador completo de Auditoría en Dashboard y diagrama Mermaid de README
console.log('4.19 Verificando apiGetAuditJobs, controles de DashboardDialog y diagrama Mermaid...');
const savedAuditLen419 = auditRows.length;
const oldAuditDate = new Date(Date.now() - 120 * 24 * 60 * 60 * 1000);
auditRows.push(
  [oldAuditDate, 'JOB-OLD-120', 'historico@centro.edu', 'Drive', 'guia.\u200Bmd', 'md', 'TEXTO', 4, 18, 850, 'OK', '', ''],
  [new Date(), 'JOB-REJ-EMPTYFAM', 'alumno@centro.edu', 'Local', 'bloqueado.exe', 'exe', '', 12, 0, 120, 'RECHAZADO', 'USER_BLOCKED', '']
);

// Seguridad RPC apiGetAuditJobs
activeEmail = 'docente@centro.edu';
const deniedJobsRes = sandbox.apiGetAuditJobs(trigTok);
assert.strictEqual(deniedJobsRes.success, false);
assert.strictEqual(deniedJobsRes.code, 'UNAUTHORIZED_ADMIN');

activeEmail = 'admin@centro.edu';
const badTokJobsRes = sandbox.apiGetAuditJobs('');
assert.strictEqual(badTokJobsRes.success, false);
assert.strictEqual(badTokJobsRes.code, 'INVALID_DIALOG_TOKEN');

const dash30Res = sandbox.apiGetDashboardAnalytics(30, sandbox.createAdminDialogToken_('admin@centro.edu'));
assert.strictEqual(dash30Res.success, true);
assert(!dash30Res.data.recentJobs.some((j) => j.jobId === 'JOB-OLD-120'), 'apiGetDashboardAnalytics(30) no debe incluir trabajos de hace 120 días');

const allJobsRes = sandbox.apiGetAuditJobs(sandbox.createAdminDialogToken_('admin@centro.edu'));
assert.strictEqual(allJobsRes.success, true);
assert.strictEqual(allJobsRes.data.totalRows, auditRows.length);
const oldJobOut = allJobsRes.data.jobs.find((j) => j.jobId === 'JOB-OLD-120');
assert(oldJobOut, 'apiGetAuditJobs debe incluir trabajos de hace 120 días aunque el periodo KPI sea 30 días');
assert.strictEqual(oldJobOut.file, 'guia.md', 'apiGetAuditJobs debe limpiar \\u200B del nombre de archivo');
assert(/^\d{2}\/\d{2}\/\d{4}\s\d{2}:\d{2}$/.test(oldJobOut.timestampFull), 'apiGetAuditJobs debe incluir timestampFull con año');
assert.strictEqual(typeof oldJobOut.tsMs, 'number');
const rejEmptyFamOut = allJobsRes.data.jobs.find((j) => j.jobId === 'JOB-REJ-EMPTYFAM');
assert(rejEmptyFamOut && rejEmptyFamOut.family === '' && rejEmptyFamOut.status === 'RECHAZADO');
for (let idx = 1; idx < allJobsRes.data.jobs.length; idx++) {
  assert(allJobsRes.data.jobs[idx - 1].tsMs >= allJobsRes.data.jobs[idx].tsMs, 'apiGetAuditJobs debe devolver los trabajos ordenados descendentemente');
}

// Verificaciones estáticas del explorador completo en DashboardDialog.html
assert(!dashLatestHtml.includes('id="showAllJobsChk" checked'), 'showAllJobsChk debe estar desactivado por defecto');
for (const reqToken of [
  'id="showAllJobsChk"',
  'id="jobsSearchInput"',
  'id="jobsUserSelect"',
  'id="jobsFamilySelect"',
  'id="jobsStatusSelect"',
  'id="jobsOriginSelect"',
  'id="jobsPageSizeSelect"',
  'id="jobsClearFiltersBtn"',
  'id="jobsExportCsvBtn"',
  'role="region"',
  'tabindex="0"',
  'aria-labelledby="jobsPanelTitle"',
  'border-collapse: separate;',
  'accent-color:',
  '@supports not (accent-color:',
  '<th scope="col">',
  'user-row-btn',
  '.apiGetAuditJobs(DIALOG_TOKEN)'
]) {
  assert(dashLatestHtml.includes(reqToken), `DashboardDialog.html debe incluir ${reqToken}`);
}

// Verificación estática de leyenda en AccessCheckDialog y consistencia/Mermaid en src/README.md
const accessCheckHtml = fs.readFileSync(path.join(srcDir, 'AccessCheckDialog.html'), 'utf8');
for (const legendTok of ['id="legendHeading"', 'g:grupo@dominio', 'Autodetección', 'Fail-Closed']) {
  assert(accessCheckHtml.includes(legendTok), `AccessCheckDialog.html debe incluir leyenda (${legendTok})`);
}

const readmeSrc = fs.readFileSync(path.join(srcDir, 'README.md'), 'utf8');
assert(readmeSrc.includes('Trigger diario de mantenimiento'), 'src/README.md debe citar el comando actual Trigger diario de mantenimiento');
assert(!readmeSrc.includes('Instalar trigger diario de limpieza'), 'src/README.md no debe mantener el nombre antiguo Instalar trigger diario de limpieza');
assert(readmeSrc.includes('wrappingWidth: 360'), 'src/README.md debe configurar wrappingWidth: 360 en Mermaid');
for (const badgeFile of ['apps-script.svg', 'drive-api.svg', 'license.svg', 'version.svg']) {
  assert(readmeSrc.includes(badgeFile), `src/README.md debe listar ${badgeFile} en la estructura del repositorio`);
}
for (const mermaidLine of [
  'BL_GRP["Evaluar grupos<br/>',
  'BLACKLIST_GROUP_ERROR"]',
  'WHITELIST_GROUP_ERROR"])',
  'Denegado<br/>NOT_IN_WHITELIST"])'
]) {
  assert(readmeSrc.includes(mermaidLine), `src/README.md debe dividir etiquetas largas en Mermaid (${mermaidLine})`);
}
auditRows.length = savedAuditLen419;

// 4.20 Edición en línea y verificación de listas de acceso + spinner/shimmer en DashboardDialog
console.log('4.20 Verificando apiSaveAccessLists, IME/aria-invalid y spinner de DashboardDialog...');
const prevActiveEmail420 = activeEmail;
resetGroupMocks();
mockGroups['claustro@centro.edu'] = { members: new Set(['docente@centro.edu']) };
mockGroups['exalumnos@centro.edu'] = { members: new Set(['antiguo@centro.edu']) };

// Bloqueo a no administradores y tokens inválidos
activeEmail = 'docente@centro.edu';
const deniedSaveRes = sandbox.apiSaveAccessLists({ verifyOnly: true }, sandbox.createAdminDialogToken_('admin@centro.edu'));
assert.strictEqual(deniedSaveRes.success, false);
assert.strictEqual(deniedSaveRes.code, 'UNAUTHORIZED_ADMIN');

activeEmail = 'admin@centro.edu';
const saveTok = sandbox.createAdminDialogToken_('admin@centro.edu');
assert.strictEqual(sandbox.apiSaveAccessLists({ verifyOnly: true }, '').code, 'INVALID_DIALOG_TOKEN');

// Rechazo de correo inválido e intento de Formula Injection sin modificar Ajustes
delete ajustesMetrics.cellUpdates['9,2'];
delete ajustesMetrics.cellUpdates['10,2'];
delete ajustesMetrics.singleUpdates['9,2'];
delete ajustesMetrics.singleUpdates['10,2'];

for (const badInput of ['=cmd@centro.edu', 'no-es-correo']) {
  const invalidSaveRes = sandbox.apiSaveAccessLists({
    blacklist: [badInput],
    whitelist: []
  }, saveTok);
  assert.strictEqual(invalidSaveRes.success, false);
  assert.strictEqual(invalidSaveRes.code, 'INVALID_ACCESS_ENTRY', `Debe rechazar entrada inválida: ${badInput}`);
  assert.strictEqual(ajustesMetrics.cellUpdates['9,2'], undefined, 'No debe modificar Ajustes ante entrada inválida');
}

// Guardado válido: deduplicación manteniendo g: sobre entrada sin prefijo + conversión de grupo:
const validSaveRes = sandbox.apiSaveAccessLists({
  blacklist: ['bloqueado@centro.edu', 'g: bloqueado@centro.edu', 'grupo:exalumnos@centro.edu'],
  whitelist: ['g:claustro@centro.edu']
}, saveTok);
assert.strictEqual(validSaveRes.success, true);
assert.strictEqual(validSaveRes.saved, true);
assert.strictEqual(JSON.stringify(validSaveRes.savedBlacklist), JSON.stringify(['g:bloqueado@centro.edu', 'g:exalumnos@centro.edu']));
assert.strictEqual(JSON.stringify(validSaveRes.savedWhitelist), JSON.stringify(['g:claustro@centro.edu']));
assert.strictEqual(ajustesMetrics.cellUpdates['9,2'], 'g:bloqueado@centro.edu, g:exalumnos@centro.edu');
assert.strictEqual(ajustesMetrics.cellUpdates['10,2'], 'g:claustro@centro.edu');
assert(validSaveRes.report && validSaveRes.report.blacklist.some((e) => e.raw === 'g:exalumnos@centro.edu' && e.kind === 'GROUP_OK'));

// verifyOnly: true re-verifica sin reescribir
delete ajustesMetrics.cellUpdates['9,2'];
delete ajustesMetrics.cellUpdates['10,2'];
const verifyOnlyRes = sandbox.apiSaveAccessLists({ verifyOnly: true }, saveTok);
assert.strictEqual(verifyOnlyRes.success, true);
assert.strictEqual(verifyOnlyRes.saved, false);
assert.strictEqual(ajustesMetrics.cellUpdates['9,2'], undefined, 'verifyOnly no debe reescribir Ajustes');
assert.strictEqual(ajustesMetrics.cellUpdates['10,2'], undefined, 'verifyOnly no debe reescribir Ajustes');
assert(verifyOnlyRes.report && verifyOnlyRes.report.whitelist.some((e) => e.raw === 'g:claustro@centro.edu' && e.kind === 'GROUP_OK'));

// Comprobaciones estáticas: spinner/shimmer de carga en DashboardDialog, editor/IME/ARIA en AccessCheckDialog y menú en Setup.gs
for (const token of [
  'id="jobsLoadingSpinner"',
  'class="loading-spinner"',
  'class="skel-cell"',
  "aria-busy', 'true'",
  'forced-color-adjust: none',
  'prefers-reduced-motion: reduce'
]) {
  assert(dashLatestHtml.includes(token), `DashboardDialog.html debe incluir ${token}`);
}
for (const token of [
  'apiSaveAccessLists',
  'event.isComposing || event.keyCode === 229',
  'aria-invalid',
  'aria-describedby',
  'id="saveVerifyBtn"',
  'id="discardBtn"'
]) {
  assert(accessCheckHtml.includes(token), `AccessCheckDialog.html debe incluir ${token}`);
}
assert(setupSrc.includes("'🛡️ Gestionar listas de acceso y grupos'"), 'Setup.gs debe renombrar el menú a Gestionar listas de acceso y grupos');

delete ajustesMetrics.cellUpdates['9,2'];
delete ajustesMetrics.cellUpdates['10,2'];
delete ajustesMetrics.singleUpdates['9,2'];
delete ajustesMetrics.singleUpdates['10,2'];
activeEmail = prevActiveEmail420;

// Restaurar estado previo
customBlacklistValue = 'bloqueado@centro.edu';
customWhitelistValue = '';
auditRows.length = savedAuditLen417;
activeEmail = prevActiveEmail417;
memCache.clear();

console.log('   OK: Todas las pruebas funcionales, de autorización, anti-Confused Deputy, Fail-Closed y anti-secuestro de sesión pasaron.');
console.log('\n✅ ENDURECIMIENTO Y VERIFICACIÓN COMPLETADOS CON ÉXITO.');
