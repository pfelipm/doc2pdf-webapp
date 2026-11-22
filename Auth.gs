/**
 * Auth.gs
 * Validación estricta (Fail-Closed) de identidad Workspace, control de acceso
 * por dominio/listas, cuota diaria, autorización de administración y tokens de diálogo.
 */

const ADMIN_LIST_CACHE_KEY_ = 'DOC2PDF_ADMINS_V1';
const DIALOG_TOKEN_TTL_SEC_ = 1800;

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
 * Devuelve la clave de fecha de hoy en la zona horaria del script (yyyy-MM-dd).
 */
function getTodayDateKey_() {
  return Utilities.formatDate(new Date(), 'Europe/Madrid', 'yyyy-MM-dd');
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
      const todayKey = getTodayDateKey_();
      const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 11).getValues();
      for (let i = rows.length - 1; i >= 0; i--) {
        const ts = rows[i][0];
        if (!(ts instanceof Date)) continue;
        const rowDay = Utilities.formatDate(ts, 'Europe/Madrid', 'yyyy-MM-dd');
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
      cache.put(cacheKey, String(count), 86400);
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
    cache.put(cacheKey, String(current + 1), 86400);
  } catch (e) {
    console.warn('Error incrementando cuota diaria:', e.message);
  }
}

/**
 * Evalúa la identidad (Fail-Closed), permisos de acceso, bypass admin en mantenimiento
 * y cuota restante del usuario activo.
 * @param {Object=} config
 * @return {Object}
 */
function getUserContext_(config) {
  const cfg = config || getConfig_();
  const activeEmail = getRealActiveEmail_();
  let effectiveEmail = '';

  try {
    effectiveEmail = (Session.getEffectiveUser().getEmail() || '').trim().toLowerCase();
  } catch (e) {
    effectiveEmail = '';
  }

  const dailyLimit = cfg.LIMITE_DIARIO_USUARIO;
  const isAdmin = isSpreadsheetAdminEmail_(activeEmail);
  const baseContext = {
    allowed: false,
    code: 'NO_SESSION_EMAIL',
    reason: '',
    email: activeEmail,
    displayEmail: activeEmail || '',
    domain: activeEmail && activeEmail.indexOf('@') !== -1 ? activeEmail.split('@')[1] : '',
    isAdmin: isAdmin,
    maintenanceBanner: '',
    dailyUsed: 0,
    dailyLimit: dailyLimit,
    remainingToday: dailyLimit > 0 ? dailyLimit : null,
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

  if (blacklist.indexOf(activeEmail) !== -1) {
    baseContext.code = 'USER_BLOCKED';
    baseContext.reason = 'Tu cuenta no tiene acceso habilitado a este servicio. Contacta con administración.';
    return baseContext;
  }

  if (whitelist.length > 0 && !isAdmin && whitelist.indexOf(activeEmail) === -1) {
    baseContext.code = 'NOT_IN_WHITELIST';
    baseContext.reason = 'El acceso está restringido a cuentas autorizadas del centro.';
    return baseContext;
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
  baseContext.remainingToday = dailyLimit > 0 ? Math.max(0, dailyLimit - used) : null;

  if (dailyLimit > 0 && used >= dailyLimit) {
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
      ctx.displayEmail || 'anonimo',
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
