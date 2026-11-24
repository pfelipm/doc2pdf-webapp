/**
 * Auth.gs
 * Validación estricta (Fail-Closed) de identidad Workspace, control de acceso
 * por dominio/listas, cuota diaria, autorización de administración y tokens de diálogo.
 */

const ADMIN_LIST_CACHE_KEY_ = 'DOC2PDF_ADMINS_V1';
const DIALOG_TOKEN_TTL_SEC_ = 1800;
const QUOTA_CACHE_TTL_SEC_ = 21600; // Máximo permitido por CacheService.put() en Apps Script (6 horas)
const GROUP_CACHE_TTL_SEC_ = 300; // 5 min, coherente con CONFIG_CACHE_TTL_SEC_
const GROUP_CACHE_EPOCH_KEY_ = 'DOC2PDF_GRP_EPOCH_V1';

// SOLO DESARROLLO / QA: deben permanecer desactivadas en producción y Git.
// Solo surten efecto cuando el usuario real es la cuenta desplegadora (real === effective).
const DEV_IMPERSONATE_ENABLED_ = false;
const DEV_IMPERSONATE_EMAIL_ = '';

/**
 * Obtiene el correo real del usuario activo autenticado en Google Workspace.
 * @return {string}
 */
function getRealActiveEmail_() {
  try {
    return (Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  } catch (e) {
    return '';
  }
}

/**
 * Resuelve la identidad efectiva de la petición web aplicando, si procede,
 * la impersonación de desarrollo estrictamente restringida al desplegador real.
 * @param {string} realEmail
 * @param {string} effectiveEmail
 * @param {{enabled: boolean, email: string}=} overrideOpt Solo uso en pruebas automatizadas.
 * @return {{email: string, realEmail: string, isImpersonating: boolean, auditEmail: string}}
 */
function resolveEffectiveIdentity_(realEmail, effectiveEmail, overrideOpt) {
  const cleanReal = String(realEmail || '').trim().toLowerCase();
  const cleanEff = String(effectiveEmail || '').trim().toLowerCase();
  const cfg = overrideOpt || {
    enabled: DEV_IMPERSONATE_ENABLED_,
    email: DEV_IMPERSONATE_EMAIL_
  };
  const target = String((cfg && cfg.email) || '').trim().toLowerCase();

  const canImpersonate = Boolean(
    cfg &&
    cfg.enabled === true &&
    target &&
    target.indexOf('@') !== -1 &&
    cleanReal &&
    cleanEff &&
    cleanReal === cleanEff &&
    target !== cleanReal
  );

  if (canImpersonate) {
    return {
      email: target,
      realEmail: cleanReal,
      isImpersonating: true,
      auditEmail: target + ' [simulado por ' + cleanReal + ']'
    };
  }

  return {
    email: cleanReal,
    realEmail: cleanReal,
    isImpersonating: false,
    auditEmail: cleanReal
  };
}

/**
 * Obtiene la época actual de invalidación para la caché de Google Groups.
 * @return {string}
 */
function getGroupCacheEpoch_() {
  try {
    const cache = CacheService.getScriptCache();
    if (!cache) return '0';
    return cache.get(GROUP_CACHE_EPOCH_KEY_) || '0';
  } catch (e) {
    return '0';
  }
}

/**
 * Incrementa la época de caché de Google Groups para invalidar entradas previas al instante.
 */
function bumpGroupCacheEpoch_() {
  try {
    const cache = CacheService.getScriptCache();
    if (!cache) return;
    const prev = Number(cache.get(GROUP_CACHE_EPOCH_KEY_) || '0');
    const now = Date.now();
    const next = String(now > prev ? now : prev + 1);
    cache.put(GROUP_CACHE_EPOCH_KEY_, next, 21600);
  } catch (e) {
    // No bloquear si CacheService no está disponible
  }
}

/**
 * Descompone una entrada normalizada de lista blanca o negra en email y bandera estricta `g:`.
 * @param {string} rawEntry
 * @return {{email: string, explicitGroup: boolean, raw: string}}
 */
function parseAccessEntry_(rawEntry) {
  const cleaned = String(rawEntry || '').trim().toLowerCase();
  if (!cleaned) {
    return { email: '', explicitGroup: false, raw: '' };
  }
  if (cleaned.indexOf('grupo:') === 0) {
    return {
      email: cleaned.slice(6).trim(),
      explicitGroup: true,
      raw: 'g:' + cleaned.slice(6).trim()
    };
  }
  if (cleaned.indexOf('g:') === 0) {
    return {
      email: cleaned.slice(2).trim(),
      explicitGroup: true,
      raw: cleaned
    };
  }
  return {
    email: cleaned,
    explicitGroup: false,
    raw: cleaned
  };
}

/**
 * Comprueba membresía directa de `userEmail` en un posible Google Group usando GroupsApp + ScriptCache (5 min).
 * Sigue política Fail-Closed ante errores en grupos explícitos (`g:`) o al consultar miembros sin permiso.
 * @param {string} userEmail
 * @param {{email: string, explicitGroup: boolean, raw: string}} parsedEntry
 * @param {string=} epochOpt
 * @return {{status: string, cached?: boolean, error?: string}}
 */
function checkUserInGroupEntry_(userEmail, parsedEntry, epochOpt) {
  const groupEmail = String((parsedEntry && parsedEntry.email) || '').trim().toLowerCase();
  const cleanUser = String(userEmail || '').trim().toLowerCase();
  if (!groupEmail || !cleanUser) {
    return { status: 'NOT_GROUP' };
  }

  let cache = null;
  try {
    cache = CacheService.getScriptCache();
  } catch (e) {
    cache = null;
  }

  const epoch = epochOpt || getGroupCacheEpoch_();
  const typeKey = 'DOC2PDF_GTYPE_' + epoch + '_' + shortSha256_(groupEmail);
  const memKey = 'DOC2PDF_GMEM_' + epoch + '_' + shortSha256_(groupEmail + '|' + cleanUser);

  if (cache) {
    try {
      const memVal = cache.get(memKey);
      if (memVal === '1') {
        return { status: 'MEMBER', cached: true };
      }
      if (memVal === '0') {
        return { status: 'NON_MEMBER', cached: true };
      }
      if (!parsedEntry.explicitGroup && cache.get(typeKey) === 'NOT_GROUP') {
        return { status: 'NOT_GROUP', cached: true };
      }
    } catch (e) {
      // Continuar con evaluación en vivo si la lectura de caché falla
    }
  }

  if (typeof GroupsApp === 'undefined' || !GroupsApp) {
    if (parsedEntry.explicitGroup) {
      return { status: 'ERROR', error: 'El servicio GroupsApp no está disponible.' };
    }
    return { status: 'NOT_GROUP' };
  }

  let group = null;
  try {
    group = GroupsApp.getGroupByEmail(groupEmail);
    if (!group) {
      throw new Error('Grupo no encontrado o inaccesible: ' + groupEmail);
    }
    if (cache) {
      try {
        cache.put(typeKey, 'GROUP', GROUP_CACHE_TTL_SEC_);
      } catch (e) {}
    }
  } catch (e) {
    if (!parsedEntry.explicitGroup) {
      if (cache) {
        try {
          cache.put(typeKey, 'NOT_GROUP', GROUP_CACHE_TTL_SEC_);
        } catch (cacheErr) {}
      }
      return { status: 'NOT_GROUP', error: e && e.message ? e.message : String(e) };
    }
    return {
      status: 'ERROR',
      error: e && e.message ? e.message : String(e)
    };
  }

  try {
    const isMember = Boolean(group.hasUser(cleanUser));
    if (cache) {
      try {
        cache.put(memKey, isMember ? '1' : '0', GROUP_CACHE_TTL_SEC_);
      } catch (cacheErr) {}
    }
    return { status: isMember ? 'MEMBER' : 'NON_MEMBER' };
  } catch (e) {
    // Grupo accesible pero lista de miembros restringida para la cuenta desplegadora -> Fail-Closed
    return {
      status: 'ERROR',
      error: e && e.message ? e.message : String(e)
    };
  }
}

/**
 * Evalúa si `userEmail` coincide con alguna entrada de una lista de acceso (usuario directo o grupo Google).
 * @param {string} userEmail
 * @param {Array<string>} entries
 * @return {{matched: boolean, matchType: ?string, matchedEntry: ?string, errorEntry: ?string, errorMessage: string}}
 */
function evaluateAccessListMatch_(userEmail, entries) {
  const cleanUser = String(userEmail || '').trim().toLowerCase();
  const list = Array.isArray(entries) ? entries : [];
  if (!cleanUser || list.length === 0) {
    return {
      matched: false,
      matchType: null,
      matchedEntry: null,
      errorEntry: null,
      errorMessage: ''
    };
  }

  const parsedEntries = [];
  for (let i = 0; i < list.length; i++) {
    const parsed = parseAccessEntry_(list[i]);
    if (!parsed.email) continue;
    if (!parsed.explicitGroup && parsed.email === cleanUser) {
      return {
        matched: true,
        matchType: 'DIRECT',
        matchedEntry: parsed.raw,
        errorEntry: null,
        errorMessage: ''
      };
    }
    parsedEntries.push(parsed);
  }

  const epoch = getGroupCacheEpoch_();
  let firstError = null;
  for (let i = 0; i < parsedEntries.length; i++) {
    const entry = parsedEntries[i];
    const check = checkUserInGroupEntry_(cleanUser, entry, epoch);
    if (check.status === 'MEMBER') {
      return {
        matched: true,
        matchType: 'GROUP',
        matchedEntry: entry.raw,
        errorEntry: null,
        errorMessage: ''
      };
    }
    if (check.status === 'ERROR' && !firstError) {
      firstError = {
        entry: entry.raw || entry.email,
        message: check.error || 'Error al consultar Google Group'
      };
    }
  }

  return {
    matched: false,
    matchType: null,
    matchedEntry: null,
    errorEntry: firstError ? firstError.entry : null,
    errorMessage: firstError ? firstError.message : ''
  };
}

/**
 * Inspecciona en vivo las entradas de una lista de acceso para diagnóstico en Sheets.
 * @param {Array<string>} entries
 * @param {string} probeEmail
 * @return {Array<{raw: string, email: string, explicitGroup: boolean, kind: string, detail: string}>}
 */
function inspectAccessListEntries_(entries, probeEmail) {
  const list = Array.isArray(entries) ? entries : [];
  const effectiveProbe = String(probeEmail || '').trim().toLowerCase() || 'doc2pdf-probe@invalid.local';
  const results = [];

  for (let i = 0; i < list.length; i++) {
    const parsed = parseAccessEntry_(list[i]);
    if (!parsed.email) continue;

    try {
      if (typeof GroupsApp === 'undefined' || !GroupsApp) {
        throw new Error('Servicio GroupsApp no disponible.');
      }
      const group = GroupsApp.getGroupByEmail(parsed.email);
      if (!group) {
        throw new Error('El grupo no existe o no es visible.');
      }
      try {
        group.hasUser(effectiveProbe);
        results.push({
          raw: parsed.raw,
          email: parsed.email,
          explicitGroup: parsed.explicitGroup,
          kind: 'GROUP_OK',
          detail: parsed.explicitGroup
            ? 'Google Group válido (modo estricto g:)'
            : 'Google Group válido (autodetectado)'
        });
      } catch (permErr) {
        results.push({
          raw: parsed.raw,
          email: parsed.email,
          explicitGroup: parsed.explicitGroup,
          kind: 'GROUP_PERMISSION_ERROR',
          detail: 'Grupo encontrado pero sin permiso de lectura de miembros: ' +
            (permErr && permErr.message ? permErr.message : String(permErr))
        });
      }
    } catch (lookupErr) {
      results.push({
        raw: parsed.raw,
        email: parsed.email,
        explicitGroup: parsed.explicitGroup,
        kind: parsed.explicitGroup ? 'EXPLICIT_GROUP_ERROR' : 'USER_OR_UNAVAILABLE',
        detail: parsed.explicitGroup
          ? 'Error Fail-Closed en grupo explícito: ' + (lookupErr && lookupErr.message ? lookupErr.message : String(lookupErr))
          : 'No resuelto como grupo (se evaluará como cuenta individual)'
      });
    }
  }

  return results;
}

/**
 * Ejecuta un diagnóstico completo de `LISTA_NEGRA_USUARIOS` y `LISTA_BLANCA_USUARIOS`,
 * verificando resolución de Google Groups, permisos de consulta de miembros y solapamientos.
 * @param {Object=} configOpt
 * @return {{blacklist: Array<Object>, whitelist: Array<Object>, overlaps: Array<string>, errors: number}}
 */
function inspectAccessLists_(configOpt) {
  const cfg = configOpt || getConfig_(true);
  let probeEmail = '';
  try {
    probeEmail = (Session.getEffectiveUser().getEmail() || getRealActiveEmail_() || '').trim().toLowerCase();
  } catch (e) {
    probeEmail = getRealActiveEmail_();
  }

  const blacklist = inspectAccessListEntries_(cfg.LISTA_NEGRA_USUARIOS || [], probeEmail);
  const whitelist = inspectAccessListEntries_(cfg.LISTA_BLANCA_USUARIOS || [], probeEmail);

  const blEmails = {};
  for (let i = 0; i < blacklist.length; i++) {
    blEmails[blacklist[i].email] = true;
  }

  const overlapSeen = {};
  const overlaps = [];
  for (let i = 0; i < whitelist.length; i++) {
    const email = whitelist[i].email;
    if (blEmails[email] && !overlapSeen[email]) {
      overlapSeen[email] = true;
      overlaps.push(email);
    }
  }

  let errors = 0;
  const all = blacklist.concat(whitelist);
  for (let i = 0; i < all.length; i++) {
    if (all[i].kind === 'GROUP_PERMISSION_ERROR' || all[i].kind === 'EXPLICIT_GROUP_ERROR') {
      errors++;
    }
  }

  return {
    blacklist: blacklist,
    whitelist: whitelist,
    overlaps: overlaps,
    errors: errors
  };
}

/**
 * Obtiene el listado deduplicado de correos con rol de administración sobre la hoja
 * (desplegador efectivo, propietario de la hoja y editores explícitos), cacheado 60 s.
 * @return {Array<string>}
 */
function getSpreadsheetAdminEmails_() {
  let cache = null;
  try {
    cache = CacheService.getScriptCache();
    const cached = cache ? cache.get(ADMIN_LIST_CACHE_KEY_) : null;
    if (cached) {
      return JSON.parse(cached);
    }
  } catch (e) {}

  const adminSet = {};
  try {
    const eff = (Session.getEffectiveUser().getEmail() || '').trim().toLowerCase();
    if (eff) adminSet[eff] = true;
  } catch (e) {}

  try {
    const ss = getSpreadsheet_();
    if (ss) {
      if (typeof ss.getOwner === 'function') {
        const owner = ss.getOwner();
        const ownerEmail = owner && typeof owner.getEmail === 'function'
          ? String(owner.getEmail() || '').trim().toLowerCase()
          : '';
        if (ownerEmail) adminSet[ownerEmail] = true;
      }
      if (typeof ss.getEditors === 'function') {
        const editors = ss.getEditors() || [];
        for (let i = 0; i < editors.length; i++) {
          const edEmail = editors[i] && typeof editors[i].getEmail === 'function'
            ? String(editors[i].getEmail() || '').trim().toLowerCase()
            : '';
          if (edEmail) adminSet[edEmail] = true;
        }
      }
    }
  } catch (e) {
    console.warn('No se pudo leer la lista de editores/propietario de la hoja:', e.message);
  }

  const admins = Object.keys(adminSet);
  try {
    if (cache) {
      cache.put(ADMIN_LIST_CACHE_KEY_, JSON.stringify(admins), 60);
    }
  } catch (e) {}
  return admins;
}

/**
 * Comprueba si un correo electrónico pertenece al grupo de administración de la hoja.
 */
function isSpreadsheetAdminEmail_(email) {
  const clean = String(email || '').trim().toLowerCase();
  if (!clean) return false;
  return getSpreadsheetAdminEmails_().indexOf(clean) !== -1;
}

/**
 * Registra un intento administrativo o de seguridad bloqueado en Auditoría.
 */
function recordSecurityRejection_(userEmail, originLabel, targetLabel, code, reason) {
  try {
    logAuditJob_({
      jobId: generateJobId_(),
      userEmail: userEmail || 'desconocido',
      origin: originLabel || 'Seguridad',
      fileName: targetLabel || '(operación)',
      ext: '',
      family: '',
      inputBytes: 0,
      pdfBytes: 0,
      durationMs: 0,
      status: 'RECHAZADO',
      detail: '[' + code + '] ' + reason,
      sheetOptions: ''
    });
  } catch (auditErr) {
    console.warn('No se pudo auditar el rechazo de seguridad:', auditErr.message);
  }
}

/**
 * Exige que quien invoca una operación de menú o panel administrativo sea
 * administrador/editor explícito de la hoja de cálculo.
 * @return {string} Correo del administrador verificado.
 */
function assertSpreadsheetAdmin_(actionLabel) {
  const activeEmail = getRealActiveEmail_();
  if (!activeEmail || !isSpreadsheetAdminEmail_(activeEmail)) {
    const reason = 'Operación administrativa no autorizada para la cuenta actual.';
    recordSecurityRejection_(activeEmail || 'anonimo', 'Admin', actionLabel || 'admin', 'UNAUTHORIZED_ADMIN', reason);
    const err = new Error(reason);
    err.code = 'UNAUTHORIZED_ADMIN';
    throw err;
  }
  return activeEmail;
}

/**
 * Protege runDailyMaintenance(e): solo permite ejecución desde un trigger temporal real
 * de Apps Script o invocación directa por un administrador de la hoja.
 */
function assertAdminOrTimeTrigger_(eventObj) {
  const isTimeTrigger = Boolean(
    eventObj &&
    typeof eventObj === 'object' &&
    typeof eventObj.triggerUid === 'string' &&
    eventObj.triggerUid.length > 0
  );
  if (isTimeTrigger) {
    return;
  }
  assertSpreadsheetAdmin_('runDailyMaintenance');
}

/**
 * Emite un token efímero vinculado al administrador actual cuando se abre un diálogo desde Sheets.
 */
function createAdminDialogToken_(adminEmail) {
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put(
    'DOC2PDF_DLG_' + token,
    JSON.stringify({ email: adminEmail }),
    DIALOG_TOKEN_TTL_SEC_
  );
  return token;
}

/**
 * Valida tanto el rol de administrador como el token efímero del diálogo de Google Sheets.
 */
function assertAdminDialogCall_(dialogToken, actionLabel) {
  const adminEmail = assertSpreadsheetAdmin_(actionLabel);
  const cleanToken = String(dialogToken || '').trim();
  let valid = false;

  if (cleanToken) {
    try {
      const raw = CacheService.getScriptCache().get('DOC2PDF_DLG_' + cleanToken);
      if (raw) {
        const parsed = JSON.parse(raw);
        if (parsed && parsed.email === adminEmail) {
          valid = true;
        }
      }
    } catch (e) {
      valid = false;
    }
  }

  if (!valid) {
    const reason = 'Token de sesión de diálogo administrativo ausente o caducado. Reabre el diálogo desde Google Sheets.';
    recordSecurityRejection_(adminEmail, 'Admin', actionLabel || 'dialogRpc', 'INVALID_DIALOG_TOKEN', reason);
    const err = new Error(reason);
    err.code = 'INVALID_DIALOG_TOKEN';
    throw err;
  }
  return adminEmail;
}

/**
 * Devuelve la clave de fecha de hoy en la zona horaria local de la hoja (yyyy-MM-dd).
 */
function getTodayDateKey_() {
  const tz = typeof getLocalTimeZone_ === 'function' ? getLocalTimeZone_() : 'Europe/Madrid';
  return Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
}

/**
 * Construye la clave de caché para la cuota diaria de un usuario.
 */
function getQuotaCacheKey_(email) {
  return 'DOC2PDF_QUOTA_' + getTodayDateKey_() + '_' + shortSha256_(String(email || '').toLowerCase());
}

/**
 * Calcula cuántas conversiones exitosas ('OK') lleva hoy el usuario indicado.
 * Usa CacheService y realiza una lectura inicial desde Auditoría si no está en caché.
 * @param {string} email
 * @return {number}
 */
function getUserDailyUsage_(email) {
  if (!email) return 0;
  const cleanEmail = String(email).trim().toLowerCase();
  const cacheKey = getQuotaCacheKey_(cleanEmail);
  let cache = null;

  try {
    cache = CacheService.getScriptCache();
    const cached = cache ? cache.get(cacheKey) : null;
    if (cached !== null && cached !== undefined) {
      const n = parseInt(cached, 10);
      if (!isNaN(n)) return n;
    }
  } catch (e) {
    console.warn('Error leyendo caché de cuota:', e.message);
  }

  let count = 0;
  try {
    const ss = getSpreadsheet_();
    const sheet = ss ? ss.getSheetByName(SHEET_AUDIT_NAME) : null;
    if (sheet && sheet.getLastRow() >= 2) {
      const tz = typeof getLocalTimeZone_ === 'function' ? getLocalTimeZone_(ss) : 'Europe/Madrid';
      const todayKey = getTodayDateKey_();
      const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 11).getValues();
      for (let i = rows.length - 1; i >= 0; i--) {
        const ts = rows[i][0];
        if (!(ts instanceof Date)) continue;
        const rowDay = Utilities.formatDate(ts, tz, 'yyyy-MM-dd');
        if (rowDay === todayKey) {
          const rowUser = String(rows[i][2] || '').trim().toLowerCase();
          const rowStatus = String(rows[i][10] || '').trim().toUpperCase();
          if (rowUser === cleanEmail && rowStatus === 'OK') {
            count++;
          }
        }
      }
    }
  } catch (e) {
    console.warn('Error calculando uso diario en Auditoría:', e.message);
  }

  try {
    if (cache) {
      cache.put(cacheKey, String(count), QUOTA_CACHE_TTL_SEC_);
    }
  } catch (e) {
    console.warn('Error guardando caché de cuota:', e.message);
  }

  return count;
}

/**
 * Incrementa el contador cacheado de cuota diaria tras una conversión exitosa.
 * @param {string} email
 */
function incrementUserDailyUsage_(email) {
  if (!email) return;
  const cleanEmail = String(email).trim().toLowerCase();
  const cacheKey = getQuotaCacheKey_(cleanEmail);
  try {
    const cache = CacheService.getScriptCache();
    if (!cache) return;
    const current = getUserDailyUsage_(cleanEmail);
    cache.put(cacheKey, String(current + 1), QUOTA_CACHE_TTL_SEC_);
  } catch (e) {
    console.warn('Error incrementando cuota diaria:', e.message);
  }
}

/**
 * Evalúa la identidad (Fail-Closed), permisos de acceso, bypass admin en mantenimiento
 * y cuota restante del usuario activo (con exención de cuota para administradores e impersonación).
 * @param {Object=} config
 * @param {{enabled: boolean, email: string}=} impersonationOverrideOpt Solo uso en pruebas automatizadas.
 * @return {Object}
 */
function getUserContext_(config, impersonationOverrideOpt) {
  const cfg = config || getConfig_();
  const realEmail = getRealActiveEmail_();
  let effectiveEmail = '';

  try {
    effectiveEmail = (Session.getEffectiveUser().getEmail() || '').trim().toLowerCase();
  } catch (e) {
    effectiveEmail = '';
  }

  const identity = resolveEffectiveIdentity_(realEmail, effectiveEmail, impersonationOverrideOpt);
  const activeEmail = identity.email;
  const dailyLimit = cfg.LIMITE_DIARIO_USUARIO;
  const isAdmin = isSpreadsheetAdminEmail_(activeEmail);
  const quotaExempt = Boolean(isAdmin || identity.isImpersonating);
  const impersonationBanner = identity.isImpersonating
    ? 'Modo simulación activo: viendo la WebApp como ' + activeEmail + ' (sesión real: ' + identity.realEmail + ' · sin consumo de cuota).'
    : '';

  const baseContext = {
    allowed: false,
    code: 'NO_SESSION_EMAIL',
    reason: '',
    email: activeEmail,
    realEmail: identity.realEmail,
    displayEmail: activeEmail || '',
    auditEmail: identity.auditEmail || '',
    domain: activeEmail && activeEmail.indexOf('@') !== -1 ? activeEmail.split('@')[1] : '',
    isAdmin: isAdmin,
    quotaExempt: quotaExempt,
    isImpersonating: identity.isImpersonating,
    impersonationBanner: impersonationBanner,
    maintenanceBanner: '',
    dailyUsed: 0,
    dailyLimit: dailyLimit,
    remainingToday: (dailyLimit > 0 && !isAdmin) ? dailyLimit : null,
    isServiceActive: cfg.SERVICIO_ACTIVO,
    maintenanceMessage: cfg.MENSAJE_MANTENIMIENTO,
    identityExposed: Boolean(activeEmail)
  };

  // Política Fail-Closed: sin email verificado en sesión, denegar acceso siempre
  if (!activeEmail) {
    baseContext.code = 'NO_SESSION_EMAIL';
    baseContext.reason = 'No se ha podido verificar tu identidad de correo en Google Workspace. Inicia sesión con tu cuenta institucional del centro.';
    return baseContext;
  }

  if (!cfg.SERVICIO_ACTIVO && !isAdmin) {
    baseContext.code = 'MAINTENANCE';
    baseContext.reason = cfg.MENSAJE_MANTENIMIENTO;
    return baseContext;
  }

  const allowedDomains = cfg.DOMINIOS_PERMITIDOS || [];
  const whitelist = cfg.LISTA_BLANCA_USUARIOS || [];
  const blacklist = cfg.LISTA_NEGRA_USUARIOS || [];

  const blEval = blacklist.length > 0
    ? evaluateAccessListMatch_(activeEmail, blacklist)
    : { matched: false, errorEntry: null, errorMessage: '' };

  if (blEval.matched) {
    baseContext.code = 'USER_BLOCKED';
    baseContext.reason = 'Tu cuenta no tiene acceso habilitado a este servicio. Contacta con administración.';
    return baseContext;
  }

  if (blEval.errorEntry) {
    baseContext.code = 'BLACKLIST_GROUP_ERROR';
    baseContext.reason = 'No se ha podido verificar la lista negra de acceso (' + blEval.errorEntry + '). Contacta con administración.';
    return baseContext;
  }

  if (whitelist.length > 0 && !isAdmin) {
    const wlEval = evaluateAccessListMatch_(activeEmail, whitelist);
    if (wlEval.errorEntry && !wlEval.matched) {
      baseContext.code = 'WHITELIST_GROUP_ERROR';
      baseContext.reason = 'No se ha podido verificar un grupo autorizado en la lista blanca (' + wlEval.errorEntry + '). Contacta con administración.';
      return baseContext;
    }
    if (!wlEval.matched) {
      baseContext.code = 'NOT_IN_WHITELIST';
      baseContext.reason = 'El acceso está restringido a cuentas autorizadas del centro.';
      return baseContext;
    }
  }

  const userDomain = baseContext.domain;
  if (allowedDomains.length > 0) {
    if (!isAdmin && allowedDomains.indexOf(userDomain) === -1) {
      baseContext.code = 'DOMAIN_NOT_ALLOWED';
      baseContext.reason = 'Tu dominio (' + userDomain + ') no está autorizado para usar este servicio.';
      return baseContext;
    }
  } else if (effectiveEmail && effectiveEmail.indexOf('@') !== -1) {
    const effectiveDomain = effectiveEmail.split('@')[1];
    const isConsumerGmail = effectiveDomain === 'gmail.com' || effectiveDomain === 'googlemail.com';
    if (isConsumerGmail) {
      if (activeEmail !== effectiveEmail) {
        baseContext.code = 'DOMAIN_NOT_ALLOWED';
        baseContext.reason = 'En cuentas personales sin DOMINIOS_PERMITIDOS solo se permite el acceso al propietario del despliegue.';
        return baseContext;
      }
    } else if (userDomain !== effectiveDomain) {
      baseContext.code = 'DOMAIN_NOT_ALLOWED';
      baseContext.reason = 'El servicio está limitado al dominio organizativo ' + effectiveDomain + '.';
      return baseContext;
    }
  }

  const used = getUserDailyUsage_(activeEmail);
  baseContext.dailyUsed = used;
  baseContext.remainingToday = (dailyLimit > 0 && !isAdmin) ? Math.max(0, dailyLimit - used) : null;

  if (dailyLimit > 0 && !isAdmin && used >= dailyLimit) {
    baseContext.code = 'QUOTA_EXCEEDED';
    baseContext.reason = 'Has alcanzado el límite diario de ' + dailyLimit + ' conversiones. Podrás volver a convertir mañana.';
    return baseContext;
  }

  baseContext.allowed = true;
  if (!cfg.SERVICIO_ACTIVO && isAdmin) {
    baseContext.code = 'MAINTENANCE_ADMIN_BYPASS';
    baseContext.maintenanceBanner = 'Modo mantenimiento activo: solo el equipo administrador puede utilizar la aplicación actualmente.';
  } else {
    baseContext.code = 'OK';
  }
  return baseContext;
}

/**
 * Verifica autorización y cuota antes de ejecutar un trabajo.
 * Si no está autorizado, registra el rechazo en Auditoría y lanza un error controlado.
 * @return {Object} Contexto de usuario validado.
 */
function assertAuthorizedForJob_(config, originLabel, fileNameForLog, extForLog, familyForLog) {
  const ctx = getUserContext_(config);
  if (!ctx.allowed) {
    recordSecurityRejection_(
      ctx.auditEmail || ctx.displayEmail || 'anonimo',
      originLabel || 'Local',
      fileNameForLog || '(solicitud)',
      ctx.code,
      ctx.reason
    );
    const err = new Error(ctx.reason);
    err.code = ctx.code;
    throw err;
  }
  return ctx;
}
