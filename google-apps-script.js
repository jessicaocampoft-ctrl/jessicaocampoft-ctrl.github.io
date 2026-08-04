// =============================================================
//  Cuidándote Fisioterapia â€” Apps Script Backend
//  Funciones: Reservas, Base de datos, Disponibilidad,
//             Panel Admin, Recordatorios diarios
// =============================================================

// IMPORTANTE: estas variables se leen desde PropertiesService (no estÃ¡n en cÃ³digo).
// Para configurarlas: en el editor de Apps Script â†’ Proyecto â†’ Propiedades del script â†’ agrega:
//   ADMIN_TOKEN   â†’ tu contraseÃ±a admin (ej: una cadena larga aleatoria)
//   GEMINI_API_KEY â†’ tu clave de Gemini AI Studio
var _props        = PropertiesService.getScriptProperties();
// Sin contraseÃ±a de respaldo en cÃ³digo: ADMIN_TOKEN debe existir en Propiedades del script.
var ADMIN_TOKEN   = _props.getProperty('ADMIN_TOKEN')   || '';
var GEMINI_API_KEY = _props.getProperty('GEMINI_API_KEY') || '';

// â”€â”€ SESIONES â”€â”€ token UUID almacenado en CacheService (TTL 4 horas)
function generateSessionToken() {
  var chars = 'abcdefghijklmnopqrstuvwxyz0123456789';
  var t = '';
  for (var i = 0; i < 48; i++) t += chars.charAt(Math.floor(Math.random() * chars.length));
  return t;
}
function createSession(user) {
  var token = generateSessionToken();
  CacheService.getScriptCache().put('sess_' + token, JSON.stringify(user || {id:'admin', nombre:'Administracion', rol:'Superadministradora'}), 14400); // 4 horas
  return token;
}
function validateSession(token) {
  if (!token || token.length < 20) return false;
  return !!CacheService.getScriptCache().get('sess_' + token);
}
function getSessionUser_(token) {
  if (!token || token.length < 20) return null;
  var raw = CacheService.getScriptCache().get('sess_' + token);
  if (!raw) return null;
  if (raw === '1') return {id:'admin', nombre:'Administracion', rol:'Superadministradora'};
  try { return JSON.parse(raw); } catch(e) { return {id:'admin', nombre:'Administracion', rol:'Superadministradora'}; }
}
function isAuxiliaryUser_(user) {
  var rol = ('' + (user && user.rol || '')).toLowerCase();
  return rol.indexOf('aux') > -1;
}
function requireAdminOnly_(user, action) {
  var ownerOnly = ['changePassword','automationSave','automationSetup','automationRun','automationQueueDone'];
  if (isAuxiliaryUser_(user) && ownerOnly.indexOf(action || '') > -1) {
    auditTeam_(user, 'Accion exclusiva de propietaria bloqueada', '', '', '', action || '');
    return {ok:false,error:'Acción exclusiva de la propietaria'};
  }
  return null;
}

function createProfessionalSession_(pro) {
  var token = generateSessionToken();
  var payload = {
    id: '' + pro.id,
    nombre: '' + pro.nombre,
    usuario: '' + pro.usuario,
    email: '' + pro.email,
    rol: '' + (pro.rol || 'Fisioterapeuta'),
    debeCambiarPassword: !!pro.debeCambiarPassword
  };
  CacheService.getScriptCache().put('prosess_' + token, JSON.stringify(payload), 14400);
  return token;
}
function validateProfessionalSession_(token) {
  if (!token || token.length < 20) return null;
  var raw = CacheService.getScriptCache().get('prosess_' + token);
  if (!raw) return null;
  try {
    CacheService.getScriptCache().put('prosess_' + token, raw, 14400);
    return JSON.parse(raw);
  } catch(e) {
    return null;
  }
}
function hashPassword_(password, salt) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + password, Utilities.Charset.UTF_8);
  return bytes.map(function(b) {
    var v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? '0' + v : v;
  }).join('');
}
function makeSalt_() {
  return Utilities.getUuid().replace(/-/g, '') + new Date().getTime();
}
function makeTempPassword_() {
  return 'Cuidandote-' + Math.floor(100000 + Math.random() * 900000);
}

// â”€â”€ RATE LIMITING LOGIN â”€â”€ mÃ¡x 5 intentos fallidos en 5 minutos (global)
function loginAllowed() {
  var v = CacheService.getScriptCache().get('login_fails');
  return !v || parseInt(v, 10) < 5;
}
function recordLoginFail() {
  var cache = CacheService.getScriptCache();
  var count = parseInt(cache.get('login_fails') || '0', 10) + 1;
  cache.put('login_fails', '' + count, 300); // ventana de 5 minutos
}
function resetLoginFails() {
  CacheService.getScriptCache().remove('login_fails');
}
var JESSICA_EMAIL = 'jessica.ocampo.ft@gmail.com';
var JESSICA_WA    = '573136467945';
var SS_NAME       = 'Citas Jessica Ocampo Fisio';

// -------------------------------------------------------------
//  GET  â€” Disponibilidad / Datos admin / Acciones admin
// -------------------------------------------------------------
function doGet(e) {
  var p = e.parameter;

  if (p.test) {
    return txt('OK - Calendario: ' + CalendarApp.getDefaultCalendar().getName());
  }

  if (p.action === 'availability' && p.date) {
    return js(getAvailability(p.date, p.service, p.modality));
  }

  // Pasaporte — lectura pública segura (requiere id + token)
  if (p.action === 'getPassportSecure') {
    return js(getPassportSecure_(p.id, p.token));
  }
  if (p.action === 'getPassport' && p.nombre) {
    return js(getLegacyPassportNotice_());
  }

  // ReseÃ±as Google â€” pÃºblico (sin token)
  if (p.action === 'getReviews') {
    return js(getGoogleReviews());
  }

  // Portal del fisioterapeuta â€” protegido por sesiÃ³n profesional.
  if (p.action === 'professionalAgenda') {
    return js(getProfessionalAgenda_(p.token));
  }

  if (!validateSession(p.token)) {
    return js({ok: false, error: 'Sin permiso'});
  }
  // Ventana deslizante: renovar TTL en cada acciÃ³n vÃ¡lida
  var sessionUser = getSessionUser_(p.token) || {id:'admin', nombre:'Administracion', rol:'Superadministradora'};
  CacheService.getScriptCache().put('sess_' + p.token, JSON.stringify(sessionUser), 14400);

  if (p.action === 'ping')          return js({ok: true});
  if (p.action === 'adminData')     return js(getAdminData());
  if (p.action === 'block')         return js(doBlock(p));
  if (p.action === 'unblock')       return js(doUnblock(p));
  if (p.action === 'updateStatus')  return js(doUpdateStatus(p, sessionUser));
  if (p.action === 'adminBook')     return js(createBooking(JSON.parse(p.data), true));
  if (p.action === 'getCalEvents')  return js(getCalendarEvents(p.from, p.to));
  if (p.action === 'cancelBooking') return js(doCancelBooking(p.id));
  if (p.action === 'editBooking')   return js(doEditBooking(JSON.parse(p.data)));
  if (p.action === 'repairRescheduledDuplicate') return js(doRepairRescheduledDuplicate(p));
  if (p.action === 'deletePatient')  return js(deletePatient(decodeURIComponent(p.nombre)));
  if (p.action === 'editPatient')    return js(editPatient(JSON.parse(p.data)));
  if (p.action === 'cleanCitasSinHora') return js(cleanCitasSinHora());
  if (p.action === 'cleanInvalidCitaTimes') return js(cleanInvalidCitaTimes());
  if (p.action === 'getReminders')   return js(getRemindersData());
  if (p.action === 'sendReminders')  return js(sendEmailReminders());
  if (p.action === 'getInactivos')   return js(getInactivosData());
  if (p.action === 'generateEval')   return js(generateEvalReport(JSON.parse(decodeURIComponent(p.data))));
  if (p.action === 'updatePago')     return js(doUpdatePago(p));
  if (p.action === 'getAdminKV')     return js(getAdminKV());
  if (p.action === 'setAdminKV')     return js(doSetAdminKV(p.data));
  if (p.action === 'generarCodigo')  return js(generarCodigo(p));
  if (p.action === 'registrarCodigo') return js(registrarCodigo(p));
  if (p.action === 'actualizarCodigo') return js(actualizarCodigo(p));
  if (p.action === 'getCodigos')     return js(getCodigos());
  if (p.action === 'crearEvento')    return js(crearEvento(p));
  if (p.action === 'eliminarEvento') return js(eliminarEvento(p));
  if (p.action === 'getEncuestaStats')    return js(getEncuestaStats_());
  if (p.action === 'autoMarcarAtendidas') return js(autoMarcarAtendidas());
  if (p.action === 'expireTemporaryReservations') return js(expireTemporaryReservations_());
  if (p.action === 'automationStatus')     return js(getAutomationStatus());
  if (p.action === 'automationSave')       { var bo0 = requireAdminOnly_(sessionUser, p.action); if (bo0) return js(bo0); return js(saveAutomationConfig(p.data)); }
  if (p.action === 'automationSetup')      { var bo1 = requireAdminOnly_(sessionUser, p.action); if (bo1) return js(bo1); return js(setupAllAutomations()); }
  if (p.action === 'automationRun')        { var bo2 = requireAdminOnly_(sessionUser, p.action); if (bo2) return js(bo2); return js(runAutomationNow(p.job || 'morning')); }
  if (p.action === 'automationQueue')      return js(getAutomationQueue(p.status || 'pending'));
  if (p.action === 'automationQueueDone')  { var bo3 = requireAdminOnly_(sessionUser, p.action); if (bo3) return js(bo3); return js(markAutomationQueueDone(p.id)); }
  if (p.action === 'getKPIHistory')        return js(getKPIHistory_());
  if (p.action === 'getWaitlist')          return js(getWaitlist());
  if (p.action === 'addWaitlist')          return js(addWaitlist(p.data));
  if (p.action === 'removeWaitlist')       return js(removeWaitlist(p.id));
  if (p.action === 'teamData')             return js(getTeamModuleData_());
  if (p.action === 'saveProfessional')     return js(saveProfessional_(p.data, sessionUser));
  if (p.action === 'resetProfessionalPassword') return js(resetProfessionalPassword_(p.id, sessionUser));
  if (p.action === 'toggleProfessional')   return js(toggleProfessional_(p.id, p.estado, sessionUser));
  if (p.action === 'deleteProfessional')   return js(deleteProfessional_(p.id, sessionUser));
  if (p.action === 'assignProfessional')   return js(assignProfessionalToAppointment_(p));
  if (p.action === 'authorizeAppointment') return js(authorizeAppointmentForProfessional_(p));
  if (p.action === 'authorizeCourtesy')     return js(authorizeCourtesy_(p, sessionUser));
  if (p.action === 'createCreditBalance')   return js(createCreditBalance_(p, sessionUser));
  if (p.action === 'applyCreditBalance')    return js(applyCreditBalance_(p, sessionUser));
  if (p.action === 'markPayablePaid')      return js(markProfessionalPayablePaid_(p.id, sessionUser));
  if (p.action === 'setupOperationsModule') return js(setupOperationsModule_());
  if (p.action === 'operationsData')        return js(getOperationsData_());
  if (p.action === 'dailyOpsData')          return js(getDailyOperationsData_(p.date, p.mode));
  if (p.action === 'saveDailyPreparation')  return js(saveDailyPreparation_(p.data, sessionUser));
  if (p.action === 'logDailyReminder')      return js(logDailyReminder_(p, sessionUser));
  if (p.action === 'addDailyObservation')   return js(addDailyObservation_(p, sessionUser));
  if (p.action === 'simulateStateMigration') return js(simulateStateMigration_());
  if (p.action === 'backupStateMigrationSheets') return js(backupStateMigrationSheets_());
  if (p.action === 'savePayment')           return js(savePayment_(p.data, sessionUser));
  if (p.action === 'verifyPayment')         return js(verifyPayment_(p, sessionUser));
  if (p.action === 'savePaymentAccount')    return js(savePaymentAccount_(p.data, sessionUser));
  if (p.action === 'passportAdminList')      return js(passportAdminList_());
  if (p.action === 'passportEnsure')         return js(passportEnsure_(decodeURIComponent(p.nombre || ''), p.telefono || '', sessionUser));
  if (p.action === 'passportSaveProgress')   return js(passportSaveProgress_(p.id, p.passport || '{}', p.descarga || '{}', sessionUser));
  if (p.action === 'passportRegenerateToken') return js(passportRegenerateToken_(p.id, sessionUser));
  if (p.action === 'passportDeactivate')     return js(passportDeactivate_(p.id, sessionUser));
  if (p.action === 'passportReactivate')     return js(passportReactivate_(p.id, sessionUser));
  if (p.action === 'passportBackupMigrate')  return js(passportBackupAndMigrate_());
  if (p.action === 'savePassport')           return js({ok:false,error:'Acción retirada: usa passportSaveProgress con sesión administrativa.'});

  return txt('Cuidándote Fisioterapia - Sistema activo');
}

// -------------------------------------------------------------
//  POST â€” Reservas de pacientes + EvaluaciÃ³n Express con fotos
// -------------------------------------------------------------
function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);
    if (d.action === 'professionalLogin') {
      return js(professionalLogin_(d.user, d.password));
    }
    if (d.action === 'professionalChangePassword') {
      return js(professionalChangePassword_(d.token, d.currentPassword, d.newPassword));
    }
    if (d.action === 'professionalMarkAttended') {
      return js(professionalMarkAttended_(d.token, d.citaId));
    }
    if (d.action === 'professionalReportIssue') {
      return js(professionalReportIssue_(d.token, d.citaId, d.tipo, d.observacion));
    }
    if (d.action === 'adminLogin') {
      if (!loginAllowed()) return js({ok: false, error: 'Demasiados intentos fallidos. Espera 5 minutos.'});
      var loginUser = ('' + (d.user || '')).trim();
      if (loginUser) {
        var pro = getProfessionalByLogin_(loginUser);
        if (!pro || pro.estado !== 'Activo' || !d.password || hashPassword_(d.password, pro.salt) !== pro.passwordHash || !isAuxiliaryUser_(pro)) {
          recordLoginFail();
          return js({ok:false,error:'Credenciales incorrectas o usuario sin permisos de auxiliar'});
        }
        resetLoginFails();
        var auxUser = {id:pro.id,nombre:pro.nombre,usuario:pro.usuario,email:pro.email,rol:pro.rol};
        var auxToken = createSession(auxUser);
        var auxData = getAdminData();
        auxData.sessionToken = auxToken;
        auxData.currentUser = auxUser;
        return js(auxData);
      }
      if (!d.password || d.password !== ADMIN_TOKEN) {
        recordLoginFail();
        return js({ok: false, error: 'Credenciales incorrectas'});
      }
      resetLoginFails();
      var sessionToken = createSession({id:'admin', nombre:'Administracion', rol:'Superadministradora'});
      var adminData = getAdminData();
      adminData.sessionToken = sessionToken;
      adminData.currentUser = {id:'admin', nombre:'Administracion', rol:'Superadministradora'};
      return js(adminData);
    }
    if (d.action === 'changePassword') {
      if (!validateSession(d.token)) return js({ok: false, error: 'Sin permiso'});
      var changeUser = getSessionUser_(d.token) || {id:'admin', nombre:'Administracion', rol:'Superadministradora'};
      var changeBlocked = requireAdminOnly_(changeUser, d.action); if (changeBlocked) return js(changeBlocked);
      if (!d.currentPassword || d.currentPassword !== ADMIN_TOKEN) return js({ok: false, error: 'La contraseÃ±a actual es incorrecta.'});
      if (!d.newPassword || d.newPassword.length < 8) return js({ok: false, error: 'La nueva contraseÃ±a debe tener al menos 8 caracteres.'});
      PropertiesService.getScriptProperties().setProperty('ADMIN_TOKEN', d.newPassword);
      return js({ok: true});
    }
    if (d.action === 'savePayment') {
      if (!validateSession(d.token)) return js({ok: false, error: 'Sin permiso'});
      var postUser = getSessionUser_(d.token) || {id:'admin', nombre:'Administracion', rol:'Superadministradora'};
      return js(savePayment_(d.data || {}, postUser));
    }
    if (d.action === 'verifyPayment') {
      if (!validateSession(d.token)) return js({ok: false, error: 'Sin permiso'});
      var postUser2 = getSessionUser_(d.token) || {id:'admin', nombre:'Administracion', rol:'Superadministradora'};
      var bp2 = requireAdminOnly_(postUser2, d.action); if (bp2) return js(bp2);
      return js(verifyPayment_(d, postUser2));
    }
    if (d.action === 'generateEval') {
      if (!validateSession(d.token)) return js({ok: false, error: 'Sin permiso'});
      return js(generateEvalReport(d.data, d.photos || {}));
    }
    return js(createBooking(d, false));
  } catch(err) {
    try { GmailApp.sendEmail(JESSICA_EMAIL, 'ERROR formulario citas', 'Error: ' + err.message + '\n\nDatos: ' + e.postData.contents); } catch(x) {}
    return js({ok: false, error: err.message});
  }
}

// -------------------------------------------------------------
//  CREAR RESERVA
// -------------------------------------------------------------
// Servicios que son solo registro de paciente â€” NO crean cita en Google Calendar
var SERVICIOS_SOLO_REGISTRO = ['Registro', 'Registro paciente', 'Registro de paciente'];

function esRegistro(servicio) {
  if (!servicio) return false;
  var s = servicio.trim().toLowerCase();
  for (var i = 0; i < SERVICIOS_SOLO_REGISTRO.length; i++) {
    if (s === SERVICIOS_SOLO_REGISTRO[i].toLowerCase()) return true;
  }
  return s.indexOf('registro') === 0; // cualquier cosa que empiece con "Registro"
}

function createBooking(d, isAdmin) {
  ensureCitasStateColumns_();
  if (isMidnightBookingTime_(d.time)) {
    return {ok: false, error: 'Ese horario es de medianoche (00:00-00:59). Para 12 del mediodia usa 12:00.'};
  }
  var scheduleCheck = isAdmin
    ? validateBookingSchedule_(d.date, d.time, d.service, d.modality)
    : validatePublicBookingSchedule_(d.date, d.time, d.service, d.modality);
  if (!scheduleCheck.ok) {
    return {ok: false, error: scheduleCheck.error};
  }
  if (!isAdmin) {
    var avail = checkAvailability(d.date, d.time, d.modality, d.service);
    if (!avail.available) return {ok: false, error: avail.reason};
  }

  var soloRegistro = esRegistro(d.service);

  // Para registros de paciente: solo guardar en hoja Pacientes (upsertPaciente ya deduplica)
  if (soloRegistro) {
    upsertPaciente(d.name, d.phone, d.email);
    return {ok: true, id: 'REG-' + new Date().getTime()};
  }

  // Lock para evitar duplicados por peticiones simultÃ¡neas (race condition)
  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch(e) { return {ok: false, error: 'Sistema ocupado, intenta de nuevo'}; }

  try {
  var price = d.priceSelected || (d.modality === 'Domicilio'
    ? d.priceD
    : (d.modality === 'Sede Campestre Recovery' ? (d.priceRecovery || d.priceP) : d.priceP));

  // Dedup: si ya existe una cita con mismo nombre+fecha+hora, devolver la existente
  var ss     = getOrCreateSheet();
  var cSheet = ss.getSheetByName('Citas');
  var cRows  = cSheet.getDataRange().getValues();
  var nameNorm = (d.name || '').toLowerCase().trim();
  for (var i = 1; i < cRows.length; i++) {
    var rowName   = ('' + (cRows[i][2]  || '')).toLowerCase().trim();
    var rowDate   = sd(cRows[i][7]);
    var rowTime   = st(cRows[i][8]);
    var rowStatus = ('' + (cRows[i][10] || '')).trim();
    if (rowName === nameNorm && rowDate === d.date && rowTime === d.time && rowStatus !== 'Cancelada') {
      return {ok: true, id: cRows[i][0]};
    }
  }

  // Crear evento en Google Calendar
  var cal   = CalendarApp.getDefaultCalendar();
  var start = parseDT(d.date, d.time);
  var mins  = getServiceDuration(d.service) + (d.modality === 'Domicilio' ? 30 : 0);
  var end   = new Date(start.getTime() + mins * 60000);
  var event = cal.createEvent('[CITA] ' + d.service + ' - ' + d.name, start, end, {
    description: buildDesc(d, price),
    location: d.modality === 'Domicilio'
      ? (d.address || 'Domicilio - direccion por confirmar')
      : (d.modality === 'Sede Campestre Recovery'
        ? 'Sede Campestre Recovery - ubicacion compartida al confirmar'
        : 'Ubicacion compartida al confirmar la reserva')
  });
  event.addEmailReminder(60);
  event.addPopupReminder(30);

  // Guardar en Google Sheets (solo citas reales)
  var clientTs = Number(d.clientTimestamp || 0);
  var id    = 'C' + (clientTs > 0 ? clientTs : new Date().getTime());
  var phoneClean = ('' + (d.phone||'')).replace(/\D/g,'');
  var rawAdminNote = '' + (d.notaAdmin || '');
  var codeNote = d.codigoReserva && rawAdminNote.indexOf(d.codigoReserva) === -1 ? '[CODIGO RESERVA: ' + d.codigoReserva + ']' : '';
  var adminNote = [rawAdminNote, codeNote].filter(Boolean).join(' ');
  cSheet.appendRow([
    id,
    new Date().toLocaleString('es-CO'),
    d.name, phoneClean, d.email,
    d.service, d.modality,
    d.date, d.time, price,
    isAdmin ? (start < new Date() ? 'Atendida' : 'Confirmada') : 'Pendiente de pago',
    d.address || '', d.notes || '', adminNote
  ]);
  var stateMap = ensureCitasStateColumns_();
  var lastRow = cSheet.getLastRow();
  var initialPaymentStatus = PAYMENT_STATUS.PENDIENTE_PAGO;
  var initialAppointmentStatus = APPOINTMENT_STATUS.RESERVADA;
  var reservationHoldMinutes = Number(operationConfigValue_('reserva_temporal_minutos', '60')) || 60;
  var reservationExpiresAt = new Date(new Date().getTime() + reservationHoldMinutes * 60000);
  cSheet.getRange(lastRow, stateMap.EstadoPago).setValue(initialPaymentStatus);
  cSheet.getRange(lastRow, stateMap.EstadoCita).setValue(initialAppointmentStatus);
  cSheet.getRange(lastRow, stateMap.EstadoMigracionOrigen).setValue(isAdmin ? 'CREADA_ADMIN' : 'RESERVA_PUBLICA');
  cSheet.getRange(lastRow, stateMap.FechaMigracionEstado).setValue(new Date());
  if (stateMap.VenceReserva) cSheet.getRange(lastRow, stateMap.VenceReserva).setValue(isAdmin ? '' : reservationExpiresAt);
  // Forzar columna Telefono como texto para evitar #ERROR! en Sheets
  cSheet.getRange(lastRow, 4).setNumberFormat('@').setValue(phoneClean);

  // Guardar/actualizar paciente en hoja Pacientes
  upsertPaciente(d.name, d.phone, d.email);

  // No enviar correos cuando la cita la crea el admin
  if (isAdmin) return {ok: true, id: id};

  // Para citas reales: enviar todos los correos y WhatsApp
  var tel  = (d.phone || '').replace(/\D/g,'');
  if (tel.length <= 10) tel = '57' + tel;
  var _waDias = ['domingo','lunes','martes','miÃ©rcoles','jueves','viernes','sÃ¡bado'];
  var _waMeses = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
  var _waDP = d.date.split('-');
  var _waFechaObj = new Date(+_waDP[0], +_waDP[1]-1, +_waDP[2]);
  var _waFecha = _waDias[_waFechaObj.getDay()] + ' ' + +_waDP[2] + ' de ' + _waMeses[+_waDP[1]-1];
  var waConfirm = 'Reserva temporal creada, ' + d.name.split(' ')[0] + '.\n\n' +
    d.service + '\n' +
    _waFecha + ' Â· ' + d.time + ' Â· ' + d.modality + '\n' +
    'Codigo de reserva: ' + (d.codigoReserva || reservationCodeFor_(id, d.date)) + '\n' +
    'Valor: ' + price + '\n\n' +
    'Para confirmar tu cita debes realizar el pago anticipado y enviar el comprobante. La cita queda autorizada solo cuando administracion confirme el pago.\n' +
    'Cuidandote Fisioterapia';
  var waLink = 'https://wa.me/' + tel + '?text=' + encodeURIComponent(waConfirm);

  GmailApp.sendEmail(
    JESSICA_EMAIL,
    'Nueva cita: ' + d.name + ' - ' + d.service + ' | ' + d.date,
    buildEmailJessica(d, price) + '\n\n>> Confirmar al paciente por WhatsApp (1 clic):\n' + waLink + '\n\nID cita: ' + id
  );

  if (d.email && d.email.indexOf('@') > 0) {
    GmailApp.sendEmail(
      d.email,
      'Reserva temporal creada - Cuidandote Fisioterapia',
      'Tu horario quedo reservado temporalmente. Para confirmar la cita debes realizar el pago anticipado y enviar el comprobante.',
      {htmlBody: buildEmailCliente(d, price), name: 'Cuidándote Fisioterapia'}
    );
  }

  return {ok: true, id: id};
  } finally {
    lock.releaseLock();
  }
}

// -------------------------------------------------------------
//  HELPERS: normaliza valores que Sheets convierte a Date/numero
// -------------------------------------------------------------
function sd(v) {
  if (!v) return '';
  if (v instanceof Date) return fmtDate(v);
  return ('' + v).split('T')[0];
}
function st(v) {
  if (!v && v !== 0) return '00:00';
  if (v instanceof Date) return pad(v.getHours()) + ':' + pad(v.getMinutes());
  if (typeof v === 'number') {
    var h = Math.floor(v * 24);
    var m = Math.round((v * 24 - h) * 60);
    return pad(h) + ':' + pad(m);
  }
  return '' + v;
}

function isMidnightBookingTime_(time) {
  var value = st(time);
  var parts = ('' + value).split(':');
  if (parts.length < 2) return false;
  var h = parseInt(parts[0], 10);
  var m = parseInt(parts[1], 10);
  return h === 0 && m >= 0 && m <= 59;
}

// -------------------------------------------------------------
//  DISPONIBILIDAD â€” lee Sheets + Calendario UNA sola vez
// -------------------------------------------------------------
function minutesFromTime_(time) {
  var t = ('' + time).split(':');
  return (parseInt(t[0], 10) || 0) * 60 + (parseInt(t[1], 10) || 0);
}

function timeFromMinutes_(mins) {
  var h = Math.floor(mins / 60), m = mins % 60;
  return pad(h) + ':' + pad(m);
}

function publicScheduleRanges_(date) {
  var dp = date.split('-');
  var d = new Date(+dp[0], +dp[1]-1, +dp[2], 0, 0, 0);
  var ranges = {
    0: [],
    1: [['08:00','16:30']],
    2: [['08:00','17:00']],
    3: [['08:00','17:00']],
    4: [['08:00','20:00']],
    5: [['08:00','20:00']],
    6: [['07:00','09:30']]
  };
  return ranges[d.getDay()] || [];
}

function clinicScheduleRanges_(date) {
  var dp = date.split('-');
  var d = new Date(+dp[0], +dp[1]-1, +dp[2], 0, 0, 0);
  var ranges = {
    0: [],
    1: [['08:00','16:30']],
    2: [['08:00','17:00']],
    3: [['08:00','17:00']],
    4: [['08:00','20:00']],
    5: [['08:00','20:00']],
    6: [['07:00','09:30'], ['14:00','18:00']]
  };
  return ranges[d.getDay()] || [];
}

function publicCandidateSlots_(date, durationMins) {
  var out = [];
  publicScheduleRanges_(date).forEach(function(range) {
    var start = minutesFromTime_(range[0]);
    var close = minutesFromTime_(range[1]);
    for (var mins = start; mins + durationMins <= close; mins += 60) {
      out.push(timeFromMinutes_(mins));
    }
  });
  return out;
}

function fitsPublicSchedule_(date, time, durationMins) {
  var start = minutesFromTime_(time);
  var end = start + durationMins;
  return publicScheduleRanges_(date).some(function(range) {
    return start >= minutesFromTime_(range[0]) && end <= minutesFromTime_(range[1]);
  });
}

function fitsClinicSchedule_(date, time, durationMins) {
  var start = minutesFromTime_(time);
  var end = start + durationMins;
  return clinicScheduleRanges_(date).some(function(range) {
    return start >= minutesFromTime_(range[0]) && end <= minutesFromTime_(range[1]);
  });
}

function validateBookingSchedule_(date, time, service, modality) {
  if (!date || !time) return {ok: false, error: 'Selecciona fecha y hora.'};
  if (isMidnightBookingTime_(time)) {
    return {ok: false, error: 'Ese horario es de medianoche (00:00-00:59). Para 12 del mediodia usa 12:00.'};
  }
  // Admin puede crear horarios manuales especiales fuera de la jornada habitual.
  // Solo se mantiene bloqueada la franja 00:00-00:59 para evitar confundir mediodia con medianoche.
  return {ok: true};
}

function validatePublicBookingSchedule_(date, time, service, modality) {
  if (!date || !time) return {ok: false, error: 'Selecciona fecha y hora.'};
  if (isMidnightBookingTime_(time)) {
    return {ok: false, error: 'Ese horario es de medianoche (00:00-00:59). Para 12 del mediodia usa 12:00.'};
  }
  var mins = getServiceDuration(service) + (modality === 'Domicilio' ? 30 : 0);
  if (!fitsClinicSchedule_(date, time, mins)) {
    return {
      ok: false,
      error: 'Ese horario no esta permitido porque la cita no cabe dentro de la jornada. El sistema no permite citas a las 9:00 p.m.; elige un horario mas temprano.'
    };
  }
  return {ok: true};
}

function validatePublicBookingSchedule_(date, time, service, modality) {
  if (!date || !time) return {ok: false, error: 'Selecciona fecha y hora.'};
  if (isMidnightBookingTime_(time)) {
    return {ok: false, error: 'Ese horario es de medianoche (00:00-00:59). Para 12 del mediodia usa 12:00.'};
  }
  var mins = getServiceDuration(service) + (modality === 'Domicilio' ? 30 : 0);
  if (!fitsPublicSchedule_(date, time, mins)) {
    return {
      ok: false,
      error: 'Ese horario no esta disponible para agenda online. Los sabados solo se agenda en la manana.'
    };
  }
  return {ok: true};
}

function getAvailability(date, service, modality) {
  var SLOTS = publicCandidateSlots_(date, getServiceDuration(service) + (modality === 'Domicilio' ? 30 : 0));
  var result = {};
  var newDur = getServiceDuration(service) + (modality === 'Domicilio' ? 30 : 0); // duraciÃ³n real del servicio que quiere agendar

  // Leer Sheets una sola vez
  var ss    = getOrCreateSheet();
  var cRows = ss.getSheetByName('Citas').getDataRange().getValues();
  var bRows = ss.getSheetByName('Bloqueos').getDataRange().getValues();

  // Leer Google Calendar una sola vez para el dia completo
  var dp = date.split('-');
  var dayStart = new Date(+dp[0], +dp[1]-1, +dp[2], 0, 0, 0);
  var dayEnd   = new Date(+dp[0], +dp[1]-1, +dp[2], 23, 59, 59);
  var calEvents = [];
  try { calEvents = CalendarApp.getDefaultCalendar().getEvents(dayStart, dayEnd); } catch(x) {}

  SLOTS.forEach(function(s) {
    var start   = parseDT(date, s);
    var endNew  = new Date(start.getTime() + newDur * 60000);
    var ok      = true;

    // 1. Google Calendar (eventos personales solamente â€” los [CITA] ya estÃ¡n en Sheets)
    for (var k = 0; k < calEvents.length && ok; k++) {
      var ev = calEvents[k];
      if (ev.isAllDayEvent()) continue;
      if (ev.getTitle().indexOf('[CITA]') === 0) continue; // ya chequeados vÃ­a Sheets
      if (start < ev.getEndTime() && endNew > ev.getStartTime()) ok = false;
    }

    // 2. Citas en Sheets â€” usar duraciÃ³n real de cada cita existente
    for (var i = 1; i < cRows.length && ok; i++) {
      var r = cRows[i];
      if (r[10] === 'Cancelada') continue;
      var rf = sd(r[7]);
      if (rf !== date) continue;
      var es  = parseDT(rf, st(r[8]));
      var em  = getServiceDuration(r[5]) + (r[6] === 'Domicilio' ? 30 : 0);
      if (start < new Date(es.getTime() + em*60000) && endNew > es) ok = false;
    }

    // 3. Bloqueos
    for (var j = 1; j < bRows.length && ok; j++) {
      var b = bRows[j];
      if (sd(b[0]) !== date) continue;
      if (start < parseDT(date, st(b[2])) && endNew > parseDT(date, st(b[1]))) ok = false;
    }

    result[s] = ok;
  });

  return {ok: true, date: date, service: service || '', modality: modality || '', duration: newDur, slots: result};
}

function checkAvailability(date, time, modality, service) {
  var start = parseDT(date, time);
  var mins  = getServiceDuration(service) + (modality === 'Domicilio' ? 30 : 0);
  var end   = new Date(start.getTime() + mins * 60000);
  var scheduleCheck = validatePublicBookingSchedule_(date, time, service, modality);
  if (!scheduleCheck.ok) {
    return {available: false, reason: scheduleCheck.error};
  }

  // 1. Google Calendar â€” bloquear solo eventos personales (no [CITA])
  try {
    var calEvents = CalendarApp.getDefaultCalendar().getEvents(start, end);
    for (var k = 0; k < calEvents.length; k++) {
      var ev = calEvents[k];
      if (ev.isAllDayEvent()) continue;
      if (ev.getTitle().indexOf('[CITA]') === 0) continue;
      return {available: false, reason: 'Ese horario no esta disponible. Por favor elige otro.'};
    }
  } catch(x) {}

  var ss = getOrCreateSheet();

  // 2. Citas existentes â€” usar duraciÃ³n real de cada servicio
  var cRows = ss.getSheetByName('Citas').getDataRange().getValues();
  for (var i = 1; i < cRows.length; i++) {
    var r = cRows[i];
    if (r[10] === 'Cancelada') continue;
    var rf = sd(r[7]);
    if (rf !== date) continue;
    var es = parseDT(rf, st(r[8]));
    var em = getServiceDuration(r[5]) + (r[6] === 'Domicilio' ? 30 : 0);
    if (start < new Date(es.getTime() + em*60000) && end > es) return {available: false, reason: 'Ese horario ya esta reservado. Por favor elige otro.'};
  }

  // 3. Bloqueos
  var bRows = ss.getSheetByName('Bloqueos').getDataRange().getValues();
  for (var j = 1; j < bRows.length; j++) {
    var b = bRows[j];
    if (sd(b[0]) !== date) continue;
    if (start < parseDT(date, st(b[2])) && end > parseDT(date, st(b[1]))) return {available: false, reason: 'Ese horario esta bloqueado.'};
  }

  return {available: true};
}

// -------------------------------------------------------------
//  ACCIONES ADMIN
// -------------------------------------------------------------
function doBlock(p) {
  var bid = 'B' + new Date().getTime();
  getOrCreateSheet().getSheetByName('Bloqueos').appendRow([
    p.date, p.startTime, p.endTime, p.reason || 'Bloqueado', 'Admin', bid
  ]);
  return {ok: true, bid: bid};
}

function doUnblock(p) {
  var sheet = getOrCreateSheet().getSheetByName('Bloqueos');
  var rows  = sheet.getDataRange().getValues();
  for (var i = rows.length - 1; i >= 1; i--) {
    // Eliminar por ID Ãºnico si existe (bloqueos nuevos)
    if (p.bid && rows[i][5] && rows[i][5] === p.bid) {
      sheet.deleteRow(i + 1);
      return {ok: true};
    }
    // Fallback para bloqueos viejos sin ID: comparar fecha + hora en mÃºltiples formatos
    if (!p.bid) {
      var rowDate  = sd(rows[i][0]);
      var rowStart = st(rows[i][1]);
      var targetTime = (p.startTime || '').trim();
      // Normalizar: si viene como decimal convertir a HH:MM
      if (!isNaN(parseFloat(targetTime)) && targetTime.indexOf(':') === -1) {
        var n = parseFloat(targetTime);
        var hh = Math.floor(n * 24);
        var mm = Math.round((n * 24 - hh) * 60);
        targetTime = (hh < 10 ? '0' : '') + hh + ':' + (mm < 10 ? '0' : '') + mm;
      }
      if (rowDate === (p.date || '').trim() && rowStart === targetTime) {
        sheet.deleteRow(i + 1);
        return {ok: true};
      }
    }
  }
  return {ok: false, error: 'No encontrado'};
}

function doUpdatePago(p) {
  var sheet = getOrCreateSheet().getSheetByName('Citas');
  var rows  = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === p.id) {
      sheet.getRange(i+1, 15).setValue(p.metodo || '');
      return {ok: true};
    }
  }
  return {ok: false, error: 'Cita no encontrada'};
}

function doUpdateStatus(p) {
  var sheet = getOrCreateSheet().getSheetByName('Citas');
  var rows  = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === p.id) {
      sheet.getRange(i+1, 11).setValue(p.status);
      if (p.note) sheet.getRange(i+1, 14).setValue(p.note);
      return {ok: true};
    }
  }
  return {ok: false, error: 'Cita no encontrada'};
}

// Devuelve eventos personales del Google Calendar (no citas) para un rango de fechas
function getCalendarEvents(from, to) {
  try {
    var dp1 = from.split('-'), dp2 = to.split('-');
    var start = new Date(+dp1[0], +dp1[1]-1, +dp1[2], 0, 0, 0);
    var end   = new Date(+dp2[0], +dp2[1]-1, +dp2[2], 23, 59, 59);
    var events = [];
    CalendarApp.getDefaultCalendar().getEvents(start, end).forEach(function(ev) {
      if (ev.getTitle().indexOf('[CITA]') === 0) return; // omitir citas del sistema
      if (ev.isAllDayEvent()) {
        events.push({title: ev.getTitle(), fecha: fmtDate(ev.getStartTime()), hora: 'Todo el dÃ­a', allDay: true});
      } else {
        events.push({
          title:   ev.getTitle(),
          fecha:   fmtDate(ev.getStartTime()),
          hora:    pad(ev.getStartTime().getHours()) + ':' + pad(ev.getStartTime().getMinutes()),
          horaFin: pad(ev.getEndTime().getHours())   + ':' + pad(ev.getEndTime().getMinutes()),
          allDay:  false
        });
      }
    });
    return {ok: true, events: events};
  } catch(x) { return {ok: false, error: x.message, events: []}; }
}

// Cancela la cita y elimina el evento del Google Calendar
function doCancelBooking(id) {
  var ss   = getOrCreateSheet();
  var rows = ss.getSheetByName('Citas').getDataRange().getValues();
  var booking = null;
  for (var i = 1; i < rows.length; i++) { if (rows[i][0] === id) { booking = rows[i]; break; } }
  var result = doUpdateStatus({id: id, status: 'Cancelada'});
  if (!result.ok) return result;
  if (booking) {
    try {
      var fecha = sd(booking[7]);
      var dp = fecha.split('-');
      var dayStart = new Date(+dp[0], +dp[1]-1, +dp[2], 0, 0, 0);
      var dayEnd   = new Date(+dp[0], +dp[1]-1, +dp[2], 23, 59, 59);
      var calEvs = CalendarApp.getDefaultCalendar().getEvents(dayStart, dayEnd);
      for (var k = 0; k < calEvs.length; k++) {
        var t = calEvs[k].getTitle() || '';
        if (t.indexOf('[CITA]') === 0 && t.indexOf(booking[2]) > -1) { calEvs[k].deleteEvent(); break; }
      }
    } catch(x) {}
  }
  var queued = 0;
  if (booking) {
    try { queued = queueWaitlistMatch_({id:id,nombre:booking[2],servicio:booking[5],fecha:sd(booking[7]),hora:st(booking[8])}); } catch(qx) {}
  }
  return {ok: true, waitlistQueued: queued};
}

// Edita una cita existente en Sheets y actualiza el evento del Calendar
function doEditBooking(d) {
  if (d.hora && isMidnightBookingTime_(d.hora)) {
    return {ok: false, error: 'Ese horario es de medianoche (00:00-00:59). Para 12 del mediodia usa 12:00.'};
  }
  var sheet = getOrCreateSheet().getSheetByName('Citas');
  var rows  = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] !== d.id) continue;
    var oldFecha = sd(rows[i][7]);
    var oldHora  = st(rows[i][8]);
    var newServicio  = d.servicio  || rows[i][5];
    var newModalidad = d.modalidad || rows[i][6];
    var newFecha     = d.fecha     || oldFecha;
    var newHora      = d.hora      || oldHora;
    var newPrecio    = d.precio    || rows[i][9];
    var scheduleCheck = validateBookingSchedule_(newFecha, newHora, newServicio, newModalidad);
    if (!scheduleCheck.ok) return {ok: false, error: scheduleCheck.error};
    if (d.servicio)           sheet.getRange(i+1, 6).setValue(d.servicio);
    if (d.modalidad)          sheet.getRange(i+1, 7).setValue(d.modalidad);
    if (d.fecha)              sheet.getRange(i+1, 8).setValue(d.fecha);
    if (d.hora)               sheet.getRange(i+1, 9).setValue(d.hora);
    if (d.precio)             sheet.getRange(i+1, 10).setValue(d.precio);
    if (d.notas !== undefined)     sheet.getRange(i+1, 13).setValue(d.notas);
    if (d.notaAdmin !== undefined) sheet.getRange(i+1, 14).setValue(d.notaAdmin);
    try {
      var dp = oldFecha.split('-');
      var dayS = new Date(+dp[0], +dp[1]-1, +dp[2], 0, 0, 0);
      var dayE = new Date(+dp[0], +dp[1]-1, +dp[2], 23, 59, 59);
      var calEvs = CalendarApp.getDefaultCalendar().getEvents(dayS, dayE);
      for (var k = 0; k < calEvs.length; k++) {
        var t = calEvs[k].getTitle() || '';
        if (t.indexOf('[CITA]') === 0 && t.indexOf(rows[i][2]) > -1) {
          var ns      = parseDT(newFecha, newHora);
          var newMins = getServiceDuration(newServicio) + (newModalidad === 'Domicilio' ? 30 : 0);
          calEvs[k].setTime(ns, new Date(ns.getTime() + newMins * 60000));
          calEvs[k].setTitle('[CITA] ' + newServicio + ' - ' + rows[i][2]);
          break;
        }
      }
    } catch(x) {}
    var dedupe = cancelDuplicateReschedules_(sheet, rows, i, {
      id: d.id,
      nombre: rows[i][2],
      servicio: newServicio,
      fecha: newFecha,
      hora: newHora,
      precio: newPrecio
    });
    if (oldFecha !== newFecha || oldHora !== newHora) {
      var map = ensureCitasStateColumns_();
      var assignment = getAssignmentMap_()[d.id];
      var state = getCitaStateFromRow_(rows[i], map);
      var nextOperational = APPOINTMENT_STATUS.RESERVADA;
      if (assignment && assignment.profesionalId && isPaymentAuthorizing_(state.estadoPago)) nextOperational = APPOINTMENT_STATUS.ASIGNADA;
      else if (isPaymentAuthorizing_(state.estadoPago)) nextOperational = APPOINTMENT_STATUS.AUTORIZADA;
      var currentRep = Number(rows[i][map.Reprogramaciones - 1] || 0) + 1;
      sheet.getRange(i+1, map.Reprogramaciones).setValue(currentRep);
      sheet.getRange(i+1, map.UltimaReprogramacion).setValue(oldFecha + ' ' + oldHora + ' -> ' + newFecha + ' ' + newHora);
      recordReprogrammingHistory_(d.id, oldFecha, oldHora, newFecha, newHora, d.motivo || '', {id:'admin', nombre:'Administracion', rol:'Superadministradora'}, currentRep);
      setCitaStates_(d.id, state.estadoPago, nextOperational, {id:'admin', nombre:'Administracion', rol:'Superadministradora'}, 'Reprogramacion: ' + oldFecha + ' ' + oldHora + ' -> ' + newFecha + ' ' + newHora + (d.motivo ? ' | Motivo: ' + d.motivo : ''));
    }
    return {ok: true, duplicatesCancelled: dedupe.cancelled, duplicateIds: dedupe.ids};
  }
  return {ok: false, error: 'Cita no encontrada'};
}

function normalizeBookingText_(v) {
  return ('' + (v || '')).toLowerCase().replace(/\s+/g, ' ').trim();
}

function bookingIsActive_(status, service) {
  var st = normalizeBookingText_(status);
  if (st === 'cancelada' || st === 'no asistio' || st === 'no asistiÃ³' || st === 'registro') return false;
  return normalizeBookingText_(service).indexOf('registro') !== 0;
}

function sameBookingIdentity_(aName, aService, aHour, bName, bService, bHour) {
  return normalizeBookingText_(aName) === normalizeBookingText_(bName)
    && normalizeBookingText_(aService) === normalizeBookingText_(bService)
    && st(aHour) === st(bHour);
}

function cancelDuplicateReschedules_(sheet, rows, keepIndex, keep) {
  var cancelled = 0, ids = [];
  for (var r = 1; r < rows.length; r++) {
    if (r === keepIndex) continue;
    var row = rows[r];
    if (!bookingIsActive_(row[10], row[5])) continue;
    if (!sameBookingIdentity_(row[2], row[5], row[8], keep.nombre, keep.servicio, keep.hora)) continue;
    var rowFecha = sd(row[7]);
    if (rowFecha !== keep.fecha) continue;
    sheet.getRange(r+1, 11).setValue('Cancelada');
    var note = ('' + (row[13] || '')).trim();
    var add  = '[AUTO] Duplicada por reprogramaciÃ³n. Cita activa: ' + keep.fecha + ' ' + keep.hora + ' (' + keep.id + ').';
    sheet.getRange(r+1, 14).setValue(note ? note + '\n' + add : add);
    cancelled++;
    ids.push(row[0]);
  }
  return {cancelled: cancelled, ids: ids};
}

function doRepairRescheduledDuplicate(p) {
  var nombre = normalizeBookingText_(p.nombre || '');
  if (!nombre) return {ok:false, error:'Falta nombre'};
  var keepFecha = sd(p.keepFecha || p.fecha || '');
  var keepHora  = st(p.keepHora || p.hora || '');
  var servicioFiltro = normalizeBookingText_(p.servicio || '');
  var ss = getOrCreateSheet();
  var sheet = ss.getSheetByName('Citas');
  var rows = sheet.getDataRange().getValues();
  var matches = [];
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (!bookingIsActive_(r[10], r[5])) continue;
    if (normalizeBookingText_(r[2]) !== nombre) continue;
    if (servicioFiltro && normalizeBookingText_(r[5]) !== servicioFiltro) continue;
    if (keepHora && st(r[8]) !== keepHora) continue;
    matches.push({idx:i, id:r[0], fecha:sd(r[7]), hora:st(r[8]), servicio:r[5], nombre:r[2]});
  }
  if (matches.length < 2) return {ok:true, repaired:0, reason:'No hay duplicados activos para ese paciente'};

  matches.sort(function(a,b){
    var fa = a.fecha || '0000-00-00', fb = b.fecha || '0000-00-00';
    if (fa !== fb) return fa.localeCompare(fb);
    return (a.hora || '').localeCompare(b.hora || '');
  });

  var keep = null;
  if (keepFecha) {
    for (var k = matches.length - 1; k >= 0; k--) {
      if (matches[k].fecha === keepFecha && (!keepHora || matches[k].hora === keepHora)) { keep = matches[k]; break; }
    }
  }
  if (!keep) keep = matches[matches.length - 1];

  var repaired = 0, cancelled = [];
  for (var m = 0; m < matches.length; m++) {
    var item = matches[m];
    if (item.idx === keep.idx) continue;
    var days = Math.abs((new Date(keep.fecha + 'T12:00:00') - new Date(item.fecha + 'T12:00:00')) / 86400000);
    if (days > 7) continue;
    sheet.getRange(item.idx+1, 11).setValue('Cancelada');
    var oldNote = ('' + (rows[item.idx][13] || '')).trim();
    var newNote = '[AUTO] Cancelada por reprogramaciÃ³n. Nueva cita activa: ' + keep.fecha + ' ' + keep.hora + ' (' + keep.id + ').';
    sheet.getRange(item.idx+1, 14).setValue(oldNote ? oldNote + '\n' + newNote : newNote);
    repaired++;
    cancelled.push({id:item.id, fecha:item.fecha, hora:item.hora});
  }
  return {ok:true, repaired:repaired, kept:{id:keep.id, fecha:keep.fecha, hora:keep.hora}, cancelled:cancelled};
}

// Elimina todos los registros de un paciente en Citas y en Pacientes
function deletePatient(nombre) {
  var ss   = getOrCreateSheet();
  var norm = (nombre || '').toLowerCase().trim();

  var cSheet = ss.getSheetByName('Citas');
  var cRows  = cSheet.getDataRange().getValues();
  for (var i = cRows.length - 1; i >= 1; i--) {
    if (('' + (cRows[i][2] || '')).toLowerCase().trim() === norm) cSheet.deleteRow(i + 1);
  }

  var pSheet = ss.getSheetByName('Pacientes');
  var pRows  = pSheet.getDataRange().getValues();
  for (var j = pRows.length - 1; j >= 1; j--) {
    if (('' + (pRows[j][0] || '')).toLowerCase().trim() === norm) pSheet.deleteRow(j + 1);
  }

  return {ok: true};
}

// Edita nombre, telÃ©fono y email en Citas y en Pacientes
function editPatient(d) {
  // d = {oldNombre, newNombre, telefono, email}
  var ss      = getOrCreateSheet();
  var phone   = (d.telefono || '').replace(/\D/g, '');
  var oldNorm = (d.oldNombre || '').toLowerCase().trim();
  var count   = 0;

  var cSheet = ss.getSheetByName('Citas');
  var cRows  = cSheet.getDataRange().getValues();
  for (var i = 1; i < cRows.length; i++) {
    if (('' + (cRows[i][2] || '')).toLowerCase().trim() !== oldNorm) continue;
    if (d.newNombre) cSheet.getRange(i+1, 3).setValue(d.newNombre);
    if (d.telefono !== undefined) cSheet.getRange(i+1, 4).setNumberFormat('@').setValue(phone);
    if (d.email    !== undefined) cSheet.getRange(i+1, 5).setValue(d.email);
    count++;
  }

  var pSheet = ss.getSheetByName('Pacientes');
  var pRows  = pSheet.getDataRange().getValues();
  var updatedPac = false;
  for (var j = 1; j < pRows.length; j++) {
    if (('' + (pRows[j][0] || '')).toLowerCase().trim() !== oldNorm) continue;
    if (d.newNombre) pSheet.getRange(j+1, 1).setValue(d.newNombre);
    if (d.telefono !== undefined) pSheet.getRange(j+1, 2).setNumberFormat('@').setValue(phone);
    if (d.email    !== undefined) pSheet.getRange(j+1, 3).setValue(d.email);
    updatedPac = true;
    break;
  }
  // Si no existÃ­a en Pacientes (solo en Citas), crearlo
  if (!updatedPac && d.newNombre) {
    var today = new Date().toLocaleDateString('es-CO');
    pSheet.appendRow([d.newNombre, phone, d.email || '', today, today]);
    pSheet.getRange(pSheet.getLastRow(), 2).setNumberFormat('@').setValue(phone);
  }

  return {ok: true, updated: count};
}

function getAdminData() {
  var ss = getOrCreateSheet();
  var stateMap = ensureCitasStateColumns_();

  var cRows = ss.getSheetByName('Citas').getDataRange().getValues();
  var citas = [];
  for (var i = 1; i < cRows.length; i++) {
    var r = cRows[i];
    var state = getCitaStateFromRow_(r, stateMap);
    citas.push({
      id: r[0], fechaReg: r[1],
      nombre: r[2],
      telefono: (function(){ var d=(r[3] instanceof Error||!r[3])?'':(''+r[3]).replace(/\D/g,''); return d.length>=7?d:''; })(),
      email: r[4],
      servicio: r[5], modalidad: r[6],
      fecha: (r[7] instanceof Date) ? fmtDate(r[7]) : (r[7] ? ('' + r[7]).split('T')[0] : ''),
      hora: st(r[8]),
      precio: r[9],
      estado: r[10], direccion: r[11], notas: r[12], notaAdmin: r[13],
      pago: ('' + (r[14] || '')).trim(),
      estadoPago: state.estadoPago,
      estadoCita: state.estadoCita,
      estadoPagoLabel: state.estadoPagoLabel,
      estadoCitaLabel: state.estadoCitaLabel,
      confirmadaVisual: state.confirmadaVisual,
      estadoVisual: state.confirmadaVisual ? 'Confirmada' : state.estadoCitaLabel
    });
  }

  var bRows = ss.getSheetByName('Bloqueos').getDataRange().getValues();
  var bloqueos = [];
  for (var j = 1; j < bRows.length; j++) {
    var b = bRows[j];
    bloqueos.push({
      bid:    b[5] || '',
      fecha: (b[0] instanceof Date) ? fmtDate(b[0]) : (b[0] ? ('' + b[0]).split('T')[0] : ''),
      inicio: st(b[1]),
      fin:    st(b[2]),
      motivo: b[3]
    });
  }

  var pRows = ss.getSheetByName('Pacientes').getDataRange().getValues();
  var pacientes = [];
  for (var k = 1; k < pRows.length; k++) {
    var p = pRows[k];
    var pPhone = ('' + (p[1]||'')).replace(/\D/g,'');
    pacientes.push({
      nombre:      ('' + (p[0]||'')).trim(),
      telefono:    pPhone,
      email:       ('' + (p[2]||'')).trim(),
      primeraVisita: (p[3] instanceof Date) ? fmtDate(p[3]) : ('' + (p[3]||'')),
      ultimaVisita:  (p[4] instanceof Date) ? fmtDate(p[4]) : ('' + (p[4]||''))
    });
  }

  // Hoja Codigos (referidos y bonos)
  var codigos = [];
  var coSheet = ss.getSheetByName('Codigos');
  if (coSheet) {
    var coRows = coSheet.getDataRange().getValues();
    for (var c = 1; c < coRows.length; c++) {
      var cr = coRows[c];
      codigos.push({
        codigo:      '' + (cr[0] || ''),
        tipo:        '' + (cr[1] || ''),
        paciente:    '' + (cr[2] || ''),
        telefono:    '' + (cr[3] || ''),
        referidoPor: '' + (cr[4] || ''),
        fecha:       '' + (cr[5] || ''),
        estado:      '' + (cr[6] || ''),
        codigoRef:   '' + (cr[7] || '')
      });
    }
  }

  // Hoja Eventos
  var eventos = [];
  var evSheet = ss.getSheetByName('Eventos');
  if (evSheet) {
    var evRows = evSheet.getDataRange().getValues();
    for (var ev = 1; ev < evRows.length; ev++) {
      var er = evRows[ev];
      if (!er[0]) continue;
      eventos.push({
        id:         '' + (er[0] || ''),
        titulo:     '' + (er[1] || ''),
        tipo:       '' + (er[2] || ''),
        fecha:      (er[3] instanceof Date) ? fmtDate(er[3]) : (er[3] ? ('' + er[3]).split('T')[0] : ''),
        horaInicio: st(er[4]),
        horaFin:    st(er[5]),
        duracion:   '' + (er[6] || ''),
        cobro:      '' + (er[7] || ''),
        notas:      '' + (er[8] || ''),
        _esEvento:  true
      });
    }
  }

  return {ok: true, citas: citas, bloqueos: bloqueos, pacientes: pacientes, codigos: codigos, eventos: eventos};
}

// -------------------------------------------------------------
//  PORTAL DEL FISIOTERAPEUTA / EQUIPO
// -------------------------------------------------------------
function teamSheet_(name) {
  var ss = getOrCreateSheet();
  var headers = {
    Profesionales: ['ID','Nombre','Usuario','Email','Rol','Estado','Servicios','Disponibilidad','TarifasJSON','Salt','PasswordHash','DebeCambiarPassword','Creado','Actualizado'],
    CitaEquipo: ['CitaID','ProfesionalID','EstadoAutorizacion','OverrideAtencion','Tarifa','Actualizado','AsignadoPor'],
    NovedadesProfesionales: ['ID','CitaID','ProfesionalID','Tipo','Observacion','Creado','EstadoAdmin'],
    AuditoriaEquipo: ['ID','Fecha','UsuarioID','UsuarioNombre','Rol','Accion','CitaID','EstadoAnterior','EstadoNuevo','Observaciones'],
    CuentasPorPagar: ['ID','ProfesionalID','CitaID','Servicio','Tarifa','Estado','Creado','Pagado','LiquidacionID']
  };
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1,1,1,headers[name].length).setValues([headers[name]]);
  }
  return sh;
}
function rowObj_(headers, row) {
  var o = {};
  for (var i = 0; i < headers.length; i++) o[headers[i]] = row[i];
  return o;
}
function getProfessionals_() {
  var sh = teamSheet_('Profesionales');
  var rows = sh.getDataRange().getValues(), out = [];
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    out.push({
      id: '' + rows[i][0], nombre: '' + (rows[i][1] || ''), usuario: '' + (rows[i][2] || ''),
      email: '' + (rows[i][3] || ''), rol: '' + (rows[i][4] || 'Fisioterapeuta'),
      estado: '' + (rows[i][5] || 'Activo'), servicios: '' + (rows[i][6] || ''),
      disponibilidad: '' + (rows[i][7] || ''), tarifasJSON: '' + (rows[i][8] || '{}'),
      salt: '' + (rows[i][9] || ''), passwordHash: '' + (rows[i][10] || ''),
      debeCambiarPassword: ('' + rows[i][11]).toLowerCase() === 'true' || rows[i][11] === true
    });
  }
  return out;
}
function getProfessionalByLogin_(user) {
  var u = (user || '').toLowerCase().trim(), list = getProfessionals_();
  for (var i = 0; i < list.length; i++) {
    if ((list[i].usuario || '').toLowerCase().trim() === u || (list[i].email || '').toLowerCase().trim() === u) return list[i];
  }
  return null;
}
function getProfessionalById_(id) {
  var list = getProfessionals_();
  for (var i = 0; i < list.length; i++) if (list[i].id === id) return list[i];
  return null;
}
function getAssignmentMap_() {
  var sh = teamSheet_('CitaEquipo');
  var rows = sh.getDataRange().getValues(), map = {};
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    map['' + rows[i][0]] = {
      citaId: '' + rows[i][0], profesionalId: '' + (rows[i][1] || ''),
      estadoAutorizacion: '' + (rows[i][2] || ''),
      overrideAtencion: ('' + rows[i][3]).toUpperCase() === 'SI',
      tarifa: rows[i][4], actualizado: rows[i][5], asignadoPor: rows[i][6]
    };
  }
  return map;
}
function getCitaById_(id) {
  var sh = getOrCreateSheet().getSheetByName('Citas');
  var rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if ('' + rows[i][0] === '' + id) {
      var r = rows[i];
      return {row: i + 1, raw: r, cita: {
        id: r[0], fechaReg: r[1], nombre: r[2], telefono: ('' + (r[3] || '')).replace(/\D/g,''),
        email: r[4], servicio: r[5], modalidad: r[6],
        fecha: (r[7] instanceof Date) ? fmtDate(r[7]) : (r[7] ? ('' + r[7]).split('T')[0] : ''),
        hora: st(r[8]), precio: r[9], estado: r[10], direccion: r[11], notas: r[12], notaAdmin: r[13],
        pago: ('' + (r[14] || '')).trim()
      }};
    }
  }
  return null;
}
function auditTeam_(user, action, citaId, prevState, newState, obs) {
  teamSheet_('AuditoriaEquipo').appendRow([
    'AUD-' + new Date().getTime() + '-' + Math.floor(Math.random()*999),
    new Date(), user && user.id ? user.id : '', user && user.nombre ? user.nombre : 'Sistema',
    user && user.rol ? user.rol : 'Sistema', action || '', citaId || '', prevState || '', newState || '', obs || ''
  ]);
}
function professionalLogin_(user, password) {
  if (!loginAllowed()) return {ok:false,error:'Demasiados intentos fallidos. Espera 5 minutos.'};
  var pro = getProfessionalByLogin_(user);
  if (!pro || pro.estado !== 'Activo' || !password || hashPassword_(password, pro.salt) !== pro.passwordHash) {
    recordLoginFail();
    auditTeam_({rol:'Sistema', nombre:'Sistema'}, 'Intento de acceso profesional fallido', '', '', '', user || '');
    return {ok:false,error:'Credenciales incorrectas o usuario inactivo'};
  }
  resetLoginFails();
  return {ok:true, professionalToken:createProfessionalSession_(pro), professional:{
    id:pro.id,nombre:pro.nombre,usuario:pro.usuario,email:pro.email,rol:pro.rol,debeCambiarPassword:pro.debeCambiarPassword
  }};
}
function professionalChangePassword_(token, currentPassword, newPassword) {
  var sess = validateProfessionalSession_(token);
  if (!sess) return {ok:false,error:'Sin permiso'};
  if (!newPassword || newPassword.length < 8) return {ok:false,error:'La nueva contraseÃ±a debe tener mÃ­nimo 8 caracteres.'};
  var sh = teamSheet_('Profesionales'), rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if ('' + rows[i][0] !== sess.id) continue;
    var salt = '' + rows[i][9];
    if (!currentPassword || hashPassword_(currentPassword, salt) !== rows[i][10]) return {ok:false,error:'La contraseÃ±a actual no coincide.'};
    var newSalt = makeSalt_();
    sh.getRange(i+1, 10, 1, 3).setValues([[newSalt, hashPassword_(newPassword, newSalt), false]]);
    sh.getRange(i+1, 14).setValue(new Date());
    auditTeam_(sess, 'Cambio de contraseÃ±a profesional', '', '', '', 'Cambio realizado por el profesional');
    return {ok:true};
  }
  return {ok:false,error:'Usuario no encontrado'};
}
var TEAM_OPERATIONAL_START_DATE = '2026-07-16';
function isProfessionalAppointmentAuthorized_(citaRow, assignment) {
  var estado = '' + (citaRow[10] || '');
  var aut = assignment ? ('' + (assignment.estadoAutorizacion || '')) : '';
  return estado === 'Autorizada para atender' || estado === 'SesiÃƒÂ³n atendida' || aut === 'Autorizada para atender' || aut === 'SesiÃƒÂ³n atendida';
}
function isProfessionalInactiveAppointment_(estado) {
  return ['Cancelada','Cancelada a tiempo','CancelaciÃƒÂ³n tardÃƒÂ­a','Reprogramada','No asistiÃƒÂ³','Reembolsada'].indexOf('' + estado) > -1;
}
function teamStateKey_(value) {
  var s = ('' + (value || '')).toLowerCase().trim();
  try { s = s.normalize('NFD').replace(/[\u0300-\u036f]/g, ''); } catch(e) {}
  return s;
}
function isProfessionalAppointmentAuthorized_(citaRow, assignment) {
  try {
    var state = getCitaStateFromRow_(citaRow, headerMap_(getOrCreateSheet().getSheetByName('Citas')));
    return state.estadoCita === APPOINTMENT_STATUS.AUTORIZADA || state.estadoCita === APPOINTMENT_STATUS.ASIGNADA || state.estadoCita === APPOINTMENT_STATUS.ATENDIDA;
  } catch(e) {}
  var estado = teamStateKey_(citaRow[10]);
  var aut = teamStateKey_(assignment ? assignment.estadoAutorizacion : '');
  var valid = ['autorizada para atender','sesion atendida','confirmada','pago verificado','cortesia autorizada','atendida'];
  return valid.indexOf(estado) > -1 || valid.indexOf(aut) > -1;
}
function isProfessionalInactiveAppointment_(estado) {
  var key = teamStateKey_(estado);
  return ['cancelada','cancelada a tiempo','cancelacion tardia','reprogramada','no asistio','reembolsada'].indexOf(key) > -1;
}
function canProfessionalAttend_(citaRow, assignment) {
  if (!isProfessionalAppointmentAuthorized_(citaRow, assignment)) return false;
  if (assignment && assignment.overrideAtencion) return true;
  var fecha = (citaRow[7] instanceof Date) ? fmtDate(citaRow[7]) : (citaRow[7] ? ('' + citaRow[7]).split('T')[0] : '');
  var hora = st(citaRow[8]);
  if (!fecha || !hora) return false;
  return new Date() >= parseDT(fecha, hora);
}
function getProfessionalAgenda_(token) {
  var sess = validateProfessionalSession_(token);
  if (!sess) {
    auditTeam_({rol:'Sistema', nombre:'Sistema'}, 'Acceso no autorizado al portal profesional', '', '', '', 'Token invÃ¡lido');
    return {ok:false,error:'Sin permiso'};
  }
  var assignments = getAssignmentMap_(), rows = getOrCreateSheet().getSheetByName('Citas').getDataRange().getValues(), citas = [];
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i], id = '' + r[0], a = assignments[id];
    if (!a || a.profesionalId !== sess.id) continue;
    var estado = '' + (r[10] || '');
    var autorizado = estado === 'Autorizada para atender' || estado === 'SesiÃ³n atendida' || a.estadoAutorizacion === 'Autorizada para atender' || a.estadoAutorizacion === 'SesiÃ³n atendida';
    if (!autorizado) continue;
    citas.push({
      id:id,
      fecha:(r[7] instanceof Date) ? fmtDate(r[7]) : (r[7] ? ('' + r[7]).split('T')[0] : ''),
      hora:st(r[8]), nombre:'' + (r[2] || ''), servicio:'' + (r[5] || ''),
      duracion:getServiceDuration(r[5]) + ((r[6] === 'Domicilio') ? 30 : 0),
      lugar:r[6] === 'Domicilio' ? ('' + (r[11] || 'Domicilio')) : 'Sede / presencial',
      modalidad:'' + (r[6] || ''), observaciones:[r[12], r[13]].filter(Boolean).join(' Â· '),
      estado:estado, autorizacion:a.estadoAutorizacion || estado, puedeAtender:canProfessionalAttend_(r, a)
    });
  }
  return {ok:true, professional:sess, citas:citas};
}
function getProfessionalAgenda_(token) {
  var sess = validateProfessionalSession_(token);
  if (!sess) {
    auditTeam_({rol:'Sistema', nombre:'Sistema'}, 'Acceso no autorizado al portal profesional', '', '', '', 'Token invalido');
    return {ok:false,error:'Sin permiso'};
  }
  var assignments = getAssignmentMap_();
  var rows = getOrCreateSheet().getSheetByName('Citas').getDataRange().getValues();
  var citas = [];
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    var id = '' + r[0];
    var a = assignments[id];
    if (!a || a.profesionalId !== sess.id) continue;
    var estado = '' + (r[10] || '');
    if (isProfessionalInactiveAppointment_(estado)) continue;
    var fecha = (r[7] instanceof Date) ? fmtDate(r[7]) : (r[7] ? ('' + r[7]).split('T')[0] : '');
    if (fecha && fecha < TEAM_OPERATIONAL_START_DATE) continue;
    var autorizado = isProfessionalAppointmentAuthorized_(r, a);
    citas.push({
      id:id,
      fecha:fecha,
      hora:st(r[8]),
      nombre:'' + (r[2] || ''),
      servicio:'' + (r[5] || ''),
      duracion:getServiceDuration(r[5]) + ((r[6] === 'Domicilio') ? 30 : 0),
      lugar:r[6] === 'Domicilio' ? ('' + (r[11] || 'Domicilio')) : 'Sede / presencial',
      modalidad:'' + (r[6] || ''),
      observaciones:[r[12], r[13]].filter(Boolean).join(' - '),
      estado:estado || (autorizado ? 'Autorizada para atender' : 'Asignada'),
      autorizacion:a.estadoAutorizacion || (autorizado ? 'Autorizada para atender' : 'Asignada pendiente de autorizacion'),
      asignada:true,
      autorizada:autorizado,
      puedeAtender:canProfessionalAttend_(r, a)
    });
  }
  return {ok:true, professional:sess, citas:citas};
}
function professionalMarkAttended_(token, citaId) {
  var sess = validateProfessionalSession_(token);
  if (!sess) return {ok:false,error:'Sin permiso'};
  var found = getCitaById_(citaId);
  if (!found) return {ok:false,error:'Cita no encontrada'};
  var assignment = getAssignmentMap_()[citaId];
  if (!assignment || assignment.profesionalId !== sess.id) {
    auditTeam_(sess, 'Intento de marcar cita ajena', citaId, '', '', 'Bloqueado por backend');
    return {ok:false,error:'No tienes permiso para esta cita'};
  }
  if (found.cita.estado === 'SesiÃ³n atendida') return {ok:false,error:'Esta sesiÃ³n ya fue marcada como atendida'};
  if (!canProfessionalAttend_(found.raw, assignment)) return {ok:false,error:'Solo puedes marcar la sesiÃ³n cuando llegue la fecha y hora de la cita'};
  var prev = found.cita.estado;
  getOrCreateSheet().getSheetByName('Citas').getRange(found.row, 11).setValue('SesiÃ³n atendida');
  var linkSh = teamSheet_('CitaEquipo'), links = linkSh.getDataRange().getValues();
  for (var i = 1; i < links.length; i++) {
    if ('' + links[i][0] === citaId) {
      linkSh.getRange(i+1, 3).setValue('SesiÃ³n atendida');
      linkSh.getRange(i+1, 6).setValue(new Date());
      break;
    }
  }
  ensurePayableForAppointment_(sess.id, citaId, found.cita.servicio, assignment.tarifa);
  auditTeam_(sess, 'MarcÃ³ sesiÃ³n como atendida', citaId, prev, 'SesiÃ³n atendida', 'AcciÃ³n realizada desde portal profesional');
  try { GmailApp.sendEmail(JESSICA_EMAIL, 'SesiÃ³n atendida: ' + found.cita.nombre, sess.nombre + ' marcÃ³ como atendida la cita ' + citaId + ' de ' + found.cita.nombre + '.'); } catch(e) {}
  return {ok:true};
}
function professionalMarkAttended_(token, citaId) {
  var sess = validateProfessionalSession_(token);
  if (!sess) return {ok:false,error:'Sin permiso'};
  ensureCitasStateColumns_();
  var attendedStatus = 'Sesi\u00f3n atendida';
  var found = getCitaById_(citaId);
  if (!found) return {ok:false,error:'Cita no encontrada'};
  var assignment = getAssignmentMap_()[citaId];
  if (!assignment || assignment.profesionalId !== sess.id) {
    auditTeam_(sess, 'Intento de marcar cita ajena', citaId, '', '', 'Bloqueado por backend');
    return {ok:false,error:'No tienes permiso para esta cita'};
  }
  if (teamStateKey_(found.cita.estado) === 'sesion atendida') return {ok:false,error:'Esta sesi\u00f3n ya fue marcada como atendida'};
  if (!canProfessionalAttend_(found.raw, assignment)) return {ok:false,error:'Solo puedes marcar la sesi\u00f3n cuando llegue la fecha y hora de la cita'};
  var prev = found.cita.estado;
  getOrCreateSheet().getSheetByName('Citas').getRange(found.row, 11).setValue(attendedStatus);
  setCitaStates_(citaId, null, APPOINTMENT_STATUS.ATENDIDA, {id:sess.id, nombre:sess.nombre, rol:'Fisioterapeuta'}, 'Sesion atendida desde portal profesional');
  var linkSh = teamSheet_('CitaEquipo'), links = linkSh.getDataRange().getValues();
  for (var i = 1; i < links.length; i++) {
    if ('' + links[i][0] === citaId) {
      linkSh.getRange(i+1, 3).setValue(attendedStatus);
      linkSh.getRange(i+1, 6).setValue(new Date());
      break;
    }
  }
  ensurePayableForAppointment_(sess.id, citaId, found.cita.servicio, assignment.tarifa);
  try {
    recordAppointmentStatusHistory_(citaId, prev, attendedStatus, {id:sess.id, nombre:sess.nombre, rol:'Fisioterapeuta'}, 'Sesion atendida desde portal profesional');
    upsertProfessionalSettlement_(sess.id, citaId, found.cita.servicio, assignment.tarifa, new Date(), {id:sess.id, nombre:sess.nombre, rol:'Fisioterapeuta'});
  } catch(e) {}
  auditTeam_(sess, 'Marco sesion como atendida', citaId, prev, attendedStatus, 'Accion realizada desde portal profesional');
  try { GmailApp.sendEmail(JESSICA_EMAIL, attendedStatus + ': ' + found.cita.nombre, sess.nombre + ' marco como atendida la cita ' + citaId + ' de ' + found.cita.nombre + '.'); } catch(e) {}
  return {ok:true};
}
function professionalReportIssue_(token, citaId, tipo, observacion) {
  var sess = validateProfessionalSession_(token);
  if (!sess) return {ok:false,error:'Sin permiso'};
  var assignment = getAssignmentMap_()[citaId];
  if (!assignment || assignment.profesionalId !== sess.id) {
    auditTeam_(sess, 'Intento de reportar novedad en cita ajena', citaId, '', '', 'Bloqueado por backend');
    return {ok:false,error:'No tienes permiso para esta cita'};
  }
  var id = 'NOV-' + new Date().getTime() + '-' + Math.floor(Math.random()*999);
  teamSheet_('NovedadesProfesionales').appendRow([id, citaId, sess.id, tipo || 'Otro', observacion || '', new Date(), 'Pendiente']);
  auditTeam_(sess, 'ReportÃ³ novedad', citaId, '', '', (tipo || 'Otro') + ' Â· ' + (observacion || ''));
  return {ok:true,id:id};
}
function saveProfessional_(data) {
  var p = parseOperationsPayload_(data);
  if (!p.nombre || !p.usuario) return {ok:false,error:'Nombre y usuario son obligatorios'};
  var sh = teamSheet_('Profesionales'), rows = sh.getDataRange().getValues(), now = new Date(), tempPassword = '';
  if (p.id) {
    for (var i = 1; i < rows.length; i++) {
      if ('' + rows[i][0] !== '' + p.id) continue;
      sh.getRange(i+1, 2, 1, 8).setValues([[p.nombre, p.usuario, p.email || '', p.rol || 'Fisioterapeuta', p.estado || 'Activo', p.servicios || '', p.disponibilidad || '', p.tarifasJSON || '{}']]);
      sh.getRange(i+1, 14).setValue(now);
      auditTeam_({rol:'Administrador', nombre:'AdministraciÃ³n'}, 'ActualizÃ³ profesional', '', '', '', p.nombre);
      return {ok:true,id:p.id};
    }
  }
  var id = 'PRO-' + new Date().getTime();
  tempPassword = p.password || makeTempPassword_();
  var salt = makeSalt_();
  sh.appendRow([id, p.nombre, p.usuario, p.email || '', p.rol || 'Fisioterapeuta', p.estado || 'Activo', p.servicios || '', p.disponibilidad || '', p.tarifasJSON || '{}', salt, hashPassword_(tempPassword, salt), true, now, now]);
  auditTeam_({rol:'Administrador', nombre:'AdministraciÃ³n'}, 'CreÃ³ profesional', '', '', '', p.nombre);
  return {ok:true,id:id,tempPassword:tempPassword};
}
function resetProfessionalPassword_(id) {
  var sh = teamSheet_('Profesionales'), rows = sh.getDataRange().getValues(), temp = makeTempPassword_(), salt = makeSalt_();
  for (var i = 1; i < rows.length; i++) {
    if ('' + rows[i][0] !== '' + id) continue;
    sh.getRange(i+1, 10, 1, 5).setValues([[salt, hashPassword_(temp, salt), true, rows[i][12] || new Date(), new Date()]]);
    auditTeam_({rol:'Administrador', nombre:'AdministraciÃ³n'}, 'RestableciÃ³ contraseÃ±a profesional', '', '', '', id);
    return {ok:true,tempPassword:temp};
  }
  return {ok:false,error:'Profesional no encontrado'};
}
function toggleProfessional_(id, estado) {
  var sh = teamSheet_('Profesionales'), rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if ('' + rows[i][0] !== '' + id) continue;
    sh.getRange(i+1, 6).setValue(estado || 'Inactivo');
    sh.getRange(i+1, 14).setValue(new Date());
    return {ok:true};
  }
  return {ok:false,error:'Profesional no encontrado'};
}
function deleteProfessional_(id) {
  var sh = teamSheet_('Profesionales'), rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if ('' + rows[i][0] !== '' + id) continue;
    sh.getRange(i+1, 6).setValue('Eliminado');
    sh.getRange(i+1, 14).setValue(new Date());
    auditTeam_({rol:'Administrador', nombre:'AdministraciÃ³n'}, 'EliminÃ³ profesional', '', '', '', rows[i][1] || id);
    return {ok:true};
  }
  return {ok:false,error:'Profesional no encontrado'};
}
function assignProfessionalToAppointment_(p) {
  ensureCitasStateColumns_();
  if (!p.citaId || !p.profesionalId) return {ok:false,error:'Falta cita o profesional'};
  var pro = getProfessionalById_(p.profesionalId);
  if (!pro || pro.estado !== 'Activo') return {ok:false,error:'Profesional inactivo o no encontrado'};
  var found = getCitaById_(p.citaId);
  if (!found) return {ok:false,error:'Cita no encontrada'};
  var state = getCitaStateFromRow_(found.raw, headerMap_(getOrCreateSheet().getSheetByName('Citas')));
  if (!isPaymentAuthorizing_(state.estadoPago)) return {ok:false,error:'La cita no puede asignarse porque el pago no estÃ¡ aprobado o no existe cortesÃ­a formal.'};
  var sh = teamSheet_('CitaEquipo'), rows = sh.getDataRange().getValues(), row = -1;
  for (var i = 1; i < rows.length; i++) if ('' + rows[i][0] === '' + p.citaId) row = i + 1;
  var values = [p.citaId, p.profesionalId, p.estadoAutorizacion || '', p.override === '1' ? 'SI' : '', p.tarifa || '', new Date(), 'AdministraciÃ³n'];
  if (row > 0) sh.getRange(row, 1, 1, values.length).setValues([values]);
  else sh.appendRow(values);
  auditTeam_({rol:'Administrador', nombre:'AdministraciÃ³n'}, 'AsignÃ³ cita a profesional', p.citaId, '', '', pro.nombre);
  setCitaStates_(p.citaId, null, APPOINTMENT_STATUS.ASIGNADA, {id:'admin', nombre:'Administracion', rol:'Superadministradora'}, 'Asignacion de fisioterapeuta: ' + pro.nombre);
  return {ok:true};
}
function authorizeAppointmentForProfessional_(p) {
  ensureCitasStateColumns_();
  var found = getCitaById_(p.citaId);
  if (!found) return {ok:false,error:'Cita no encontrada'};
  var assignment = getAssignmentMap_()[p.citaId];
  if (!assignment || !assignment.profesionalId) return {ok:false,error:'Primero asigna un fisioterapeuta'};
  var inactiveStates = ['Cancelada','Cancelada a tiempo','CancelaciÃ³n tardÃ­a','Reprogramada','No asistiÃ³','Reembolsada'];
  var active = inactiveStates.indexOf(found.cita.estado) === -1;
  var state = getCitaStateFromRow_(found.raw, headerMap_(getOrCreateSheet().getSheetByName('Citas')));
  var paid = isPaymentAuthorizing_(state.estadoPago);
  if (!active) return {ok:false,error:'La cita no estÃ¡ activa'};
  if (!paid && p.excepcion !== '1') return {ok:false,error:'Falta pago verificado. Usa excepciÃ³n si quieres autorizar cortesÃ­a o caso especial.'};
  var prev = found.cita.estado;
  getOrCreateSheet().getSheetByName('Citas').getRange(found.row, 11).setValue('Autorizada para atender');
  setCitaStates_(p.citaId, state.estadoPago, APPOINTMENT_STATUS.AUTORIZADA, {id:'admin', nombre:'Administracion', rol:'Superadministradora'}, 'Autorizacion administrativa para atender');
  var sh = teamSheet_('CitaEquipo'), rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if ('' + rows[i][0] === '' + p.citaId) {
      sh.getRange(i+1, 3).setValue('Autorizada para atender');
      sh.getRange(i+1, 6).setValue(new Date());
      break;
    }
  }
  auditTeam_({rol:'Administrador', nombre:'AdministraciÃ³n'}, 'AutorizÃ³ cita para atender', p.citaId, prev, 'Autorizada para atender', p.excepcion === '1' ? 'Con excepciÃ³n administrativa' : '');
  return {ok:true};
}
function ensurePayableForAppointment_(professionalId, citaId, servicio, tarifa) {
  var sh = teamSheet_('CuentasPorPagar'), rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) if ('' + rows[i][2] === citaId) return;
  sh.appendRow(['PAG-' + new Date().getTime(), professionalId, citaId, servicio || '', tarifa || '', 'Pendiente', new Date(), '', '']);
}
function markProfessionalPayablePaid_(id) {
  var sh = teamSheet_('CuentasPorPagar'), rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if ('' + rows[i][0] !== '' + id) continue;
    sh.getRange(i+1, 6, 1, 3).setValues([['Pagada', new Date(), 'LIQ-' + Utilities.formatDate(new Date(), 'America/Bogota', 'yyyyMM')]]);
    return {ok:true};
  }
  return {ok:false,error:'Cuenta no encontrada'};
}
function getTeamModuleData_() {
  function sheetRows(name) {
    var sh = teamSheet_(name), values = sh.getDataRange().getValues();
    var headers = values[0], out = [];
    for (var i = 1; i < values.length; i++) {
      if (!values[i][0]) continue;
      var o = rowObj_(headers, values[i]);
      Object.keys(o).forEach(function(k) {
        if (o[k] instanceof Date) o[k] = o[k].toISOString();
        else o[k] = '' + (o[k] || '');
      });
      out.push(o);
    }
    return out;
  }
  return {
    ok:true,
    profesionales:getProfessionals_().map(function(p){return {id:p.id,nombre:p.nombre,usuario:p.usuario,email:p.email,rol:p.rol,estado:p.estado,servicios:p.servicios,disponibilidad:p.disponibilidad,tarifasJSON:p.tarifasJSON,debeCambiarPassword:p.debeCambiarPassword};}),
    asignaciones:sheetRows('CitaEquipo'),
    novedades:sheetRows('NovedadesProfesionales'),
    auditoria:sheetRows('AuditoriaEquipo').slice(-80).reverse(),
    cuentas:sheetRows('CuentasPorPagar')
  };
}

// -------------------------------------------------------------
//  HELPERS PLANES â€” detecciÃ³n y lÃ³gica de pagos
// -------------------------------------------------------------
// -------------------------------------------------------------
//  MODULO OPERATIVO: PAGOS, PLANES, ROLES, HISTORIAL
// -------------------------------------------------------------
var APPOINTMENT_STATUS_CATALOG = [
  'Solicitud recibida','Pendiente de pago','Pago por verificar','Pago rechazado',
  'Confirmada','Pago verificado','CortesÃ­a autorizada','Autorizada para atender',
  'SesiÃ³n iniciada','SesiÃ³n atendida','Cerrada','Cancelada a tiempo',
  'CancelaciÃ³n tardÃ­a','No asistiÃ³','Reprogramada','Saldo a favor',
  'Reserva vencida','Cancelada','Atendida','Pendiente'
];

var PAYMENT_STATUS = {
  PENDIENTE_PAGO: 'PENDIENTE_PAGO',
  COMPROBANTE_RECIBIDO: 'COMPROBANTE_RECIBIDO',
  PAGO_APROBADO: 'PAGO_APROBADO',
  PAGO_RECHAZADO: 'PAGO_RECHAZADO',
  REEMBOLSADO: 'REEMBOLSADO',
  NO_REQUIERE_PAGO: 'NO_REQUIERE_PAGO'
};

var APPOINTMENT_STATUS = {
  RESERVADA: 'RESERVADA',
  AUTORIZADA: 'AUTORIZADA',
  ASIGNADA: 'ASIGNADA',
  ATENDIDA: 'ATENDIDA',
  CANCELADA: 'CANCELADA',
  NO_ASISTIO: 'NO_ASISTIO'
};

var CITA_STATE_HEADERS = [
  'EstadoPago','EstadoCita','EstadoMigracionOrigen','FechaMigracionEstado',
  'Reprogramaciones','UltimaReprogramacion','RequiereCierreAdmin','VenceReserva'
];

function normalizeKey_(v) {
  return ('' + (v || '')).toLowerCase()
    .replace(/[Ã¡Ã Ã¤Ã¢]/g,'a').replace(/[Ã©Ã¨Ã«Ãª]/g,'e').replace(/[Ã­Ã¬Ã¯Ã®]/g,'i')
    .replace(/[Ã³Ã²Ã¶Ã´]/g,'o').replace(/[ÃºÃ¹Ã¼Ã»]/g,'u').replace(/Ã±/g,'n').trim();
}

function paymentStatusLabel_(s) {
  var labels = {
    PENDIENTE_PAGO:'Pendiente de pago',
    COMPROBANTE_RECIBIDO:'Comprobante recibido',
    PAGO_APROBADO:'Pago aprobado',
    PAGO_RECHAZADO:'Pago rechazado',
    REEMBOLSADO:'Reembolsado',
    NO_REQUIERE_PAGO:'No requiere pago'
  };
  return labels[s] || s || 'Sin estado de pago';
}

function appointmentStatusLabel_(s) {
  var labels = {
    RESERVADA:'Reservada',
    AUTORIZADA:'Autorizada',
    ASIGNADA:'Asignada',
    ATENDIDA:'Atendida',
    CANCELADA:'Cancelada',
    NO_ASISTIO:'No asistiÃ³'
  };
  return labels[s] || s || 'Sin estado de cita';
}

function legacyPaymentToNew_(legacyState, pagoValue) {
  var stKey = normalizeKey_(legacyState);
  var payKey = normalizeKey_(pagoValue);
  if (stKey === 'saldo a favor') return {status:'__SALDO_A_FAVOR__', ambiguous:true, reason:'Saldo a favor debe convertirse en movimiento financiero, no en reembolso.'};
  if (stKey === 'cortesia autorizada') return {status:PAYMENT_STATUS.NO_REQUIERE_PAGO, courtesy:true};
  if (stKey === 'reembolsada') return {status:PAYMENT_STATUS.REEMBOLSADO};
  if (stKey === 'pago rechazado' || stKey === 'rechazado') return {status:PAYMENT_STATUS.PAGO_RECHAZADO};
  if (stKey === 'pago por verificar' || stKey === 'por verificar') return {status:PAYMENT_STATUS.COMPROBANTE_RECIBIDO};
  if (stKey === 'pago verificado' || stKey === 'autorizada para atender' || stKey === 'sesion atendida' || stKey === 'atendida') return {status:PAYMENT_STATUS.PAGO_APROBADO};
  if (stKey === 'confirmada' && payKey) return {status:PAYMENT_STATUS.PAGO_APROBADO};
  if (payKey && payKey !== 'pendiente') return {status:PAYMENT_STATUS.PAGO_APROBADO};
  return {status:PAYMENT_STATUS.PENDIENTE_PAGO};
}

function legacyAppointmentToNew_(legacyState, paymentStatus, hasAssignment) {
  var stKey = normalizeKey_(legacyState);
  if (stKey === 'registro') return {status:'__NO_APLICA__', ambiguous:false};
  if (stKey === 'saldo a favor') return {status:APPOINTMENT_STATUS.CANCELADA, ambiguous:true, reason:'Saldo a favor requiere revisiÃ³n financiera manual.'};
  if (stKey === 'reprogramada') return {status:'__REPROGRAMADA_OPERATIVA__', ambiguous:true, reason:'Reprogramada debe ser evento histÃ³rico; se debe definir estado operativo actual.'};
  if (stKey === 'cancelada' || stKey === 'cancelada a tiempo' || stKey === 'cancelacion tardia' || stKey === 'reserva vencida' || stKey === 'reembolsada') return {status:APPOINTMENT_STATUS.CANCELADA};
  if (stKey === 'no asistio') return {status:APPOINTMENT_STATUS.NO_ASISTIO};
  if (stKey === 'atendida' || stKey === 'sesion atendida' || stKey === 'cerrada') return {status:APPOINTMENT_STATUS.ATENDIDA};
  if (hasAssignment) return {status:APPOINTMENT_STATUS.ASIGNADA};
  if (paymentStatus === PAYMENT_STATUS.PAGO_APROBADO || paymentStatus === PAYMENT_STATUS.NO_REQUIERE_PAGO) return {status:APPOINTMENT_STATUS.AUTORIZADA};
  return {status:APPOINTMENT_STATUS.RESERVADA};
}

function isPaymentAuthorizing_(status) {
  return status === PAYMENT_STATUS.PAGO_APROBADO || status === PAYMENT_STATUS.NO_REQUIERE_PAGO;
}

function isConfirmedVisual_(estadoPago, estadoCita) {
  return estadoPago === PAYMENT_STATUS.PAGO_APROBADO && (estadoCita === APPOINTMENT_STATUS.AUTORIZADA || estadoCita === APPOINTMENT_STATUS.ASIGNADA);
}

function ensureCitasStateColumns_() {
  var sh = getOrCreateSheet().getSheetByName('Citas');
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h){ return '' + (h || ''); });
  CITA_STATE_HEADERS.forEach(function(h) {
    if (headers.indexOf(h) === -1) {
      sh.getRange(1, sh.getLastColumn() + 1).setValue(h);
      headers.push(h);
    }
  });
  return headerMap_(sh);
}

function headerMap_(sh) {
  var headers = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
  var map = {};
  for (var i = 0; i < headers.length; i++) map['' + headers[i]] = i + 1;
  return map;
}

function getCitaStateFromRow_(row, map) {
  map = map || {};
  var legacy = row[10] || '';
  var pago = map.EstadoPago ? row[map.EstadoPago - 1] : '';
  var cita = map.EstadoCita ? row[map.EstadoCita - 1] : '';
  if (!pago || !cita) {
    var pay = legacyPaymentToNew_(legacy, row[14]);
    var app = legacyAppointmentToNew_(legacy, pay.status, false);
    pago = pago || (pay.status.indexOf('__') === 0 ? PAYMENT_STATUS.PENDIENTE_PAGO : pay.status);
    cita = cita || (app.status.indexOf('__') === 0 ? APPOINTMENT_STATUS.RESERVADA : app.status);
  }
  return {
    estadoPago: '' + pago,
    estadoCita: '' + cita,
    estadoPagoLabel: paymentStatusLabel_(pago),
    estadoCitaLabel: appointmentStatusLabel_(cita),
    confirmadaVisual: isConfirmedVisual_(pago, cita)
  };
}

function setCitaStates_(citaId, estadoPago, estadoCita, user, obs) {
  var sh = getOrCreateSheet().getSheetByName('Citas');
  var map = ensureCitasStateColumns_();
  var rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if ('' + rows[i][0] !== '' + citaId) continue;
    var prev = getCitaStateFromRow_(rows[i], map);
    if (estadoPago) sh.getRange(i + 1, map.EstadoPago).setValue(estadoPago);
    if (estadoCita) sh.getRange(i + 1, map.EstadoCita).setValue(estadoCita);
    var nextRow = rows[i].slice();
    if (estadoPago) nextRow[map.EstadoPago - 1] = estadoPago;
    if (estadoCita) nextRow[map.EstadoCita - 1] = estadoCita;
    try { passportSyncAppointment_(nextRow, map, user); } catch(passportErr) {
      try { auditGeneral_(user, 'Error sincronizando Pasaporte', 'Cita', citaId, '', passportErr.message, obs || ''); } catch(auditErr) {}
    }
    recordAppointmentStatusHistory_(
      citaId,
      prev.estadoPago + ' / ' + prev.estadoCita,
      (estadoPago || prev.estadoPago) + ' / ' + (estadoCita || prev.estadoCita),
      user,
      obs || 'Actualizacion de estados normalizados'
    );
    return {ok:true};
  }
  return {ok:false,error:'Cita no encontrada'};
}

function classifyLegacyCitaForMigration_(row, assignmentMap) {
  var citaId = '' + (row[0] || '');
  var legacy = '' + (row[10] || '');
  var pagoRaw = '' + (row[14] || '');
  var hasAssignment = !!(assignmentMap && assignmentMap[citaId] && assignmentMap[citaId].profesionalId);
  var pay = legacyPaymentToNew_(legacy, pagoRaw);
  var app = legacyAppointmentToNew_(legacy, pay.status, hasAssignment);
  var ambiguous = !!(pay.ambiguous || app.ambiguous || pay.status.indexOf('__') === 0 || app.status.indexOf('__') === 0);
  return {
    id:citaId,
    estadoAntiguo:legacy,
    pagoRegistrado:pagoRaw,
    fecha:sd(row[7]),
    hora:st(row[8]),
    estadoPago:pay.status.indexOf('__') === 0 ? '' : pay.status,
    estadoCita:app.status.indexOf('__') === 0 ? '' : app.status,
    requiereRevision:ambiguous,
    motivoDuda:[pay.reason, app.reason].filter(Boolean).join(' | '),
    propuesta: (pay.status.indexOf('__') === 0 ? pay.status : pay.status) + ' / ' + (app.status.indexOf('__') === 0 ? app.status : app.status)
  };
}

function simulateStateMigration_() {
  var ss = getOrCreateSheet();
  var sh = ss.getSheetByName('Citas');
  var rows = sh.getDataRange().getValues();
  var assignments = getAssignmentMap_();
  var countsPago = {}, countsCita = {}, ambiguos = [], total = 0;
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    total++;
    var c = classifyLegacyCitaForMigration_(rows[i], assignments);
    if (c.requiereRevision) ambiguos.push(c);
    else {
      countsPago[c.estadoPago] = (countsPago[c.estadoPago] || 0) + 1;
      countsCita[c.estadoCita] = (countsCita[c.estadoCita] || 0) + 1;
    }
  }
  return {ok:true,total:total,countsPago:countsPago,countsCita:countsCita,requierenRevision:ambiguos};
}

function backupStateMigrationSheets_() {
  var ss = getOrCreateSheet();
  var stamp = Utilities.formatDate(new Date(), 'America/Bogota', 'yyyyMMdd-HHmmss');
  var copy = DriveApp.getFileById(ss.getId()).makeCopy('BACKUP estados citas - ' + stamp);
  return {ok:true,backupFileId:copy.getId(),backupName:copy.getName(),createdAt:stamp};
}

function operationsSheet_(name) {
  var ss = getOrCreateSheet();
  var headers = {
    Roles: ['ID','Nombre','Descripcion','Permisos','Estado','Creado','Actualizado'],
    UsuariosAdmin: ['ID','Nombre','Email','Rol','Estado','Creado','Actualizado'],
    CuentasPago: ['ID','Medio','Tipo','Numero','Titular','Estado','Orden','Actualizado'],
    ConfiguracionOperativa: ['Clave','Valor','Descripcion','Actualizado'],
    HistorialEstadosCita: ['ID','CitaID','CodigoReserva','EstadoAnterior','EstadoNuevo','Fecha','UsuarioID','UsuarioNombre','Rol','Observacion','TipoCambio','FechaHoraAnterior','FechaHoraNueva','Motivo','Reprogramaciones'],
    Pagos: ['ID','CodigoReserva','CitaID','Cliente','ServicioPlan','ValorEsperado','ValorRecibido','MedioPago','CuentaReceptora','FechaPago','FechaVerificacion','Comprobante','EstadoPago','UsuarioVerifico','Observaciones','CuotaNumero','SaldoPendiente','Creado','Actualizado'],
    ComprobantesPago: ['ID','PagoID','CodigoReserva','CitaID','NombreArchivo','TipoArchivo','Tamano','DriveFileID','Estado','Creado','Observaciones','Hash'],
    SaldosFavor: ['ID','Cliente','Telefono','ValorDisponible','CitaOrigenID','PagoOrigenID','FechaCreacion','FechaVencimiento','ValorUtilizado','SaldoRestante','UsuarioAutorizo','Estado','Observaciones','Actualizado'],
    Cortesias: ['ID','CitaID','Cliente','Servicio','ValorComercial','Motivo','UsuarioAutorizo','FechaAutorizacion','Estado','Observaciones'],
    AgendaPreparacion: ['ID','FechaObjetivo','Tipo','EstadoPreparacion','UsuarioID','UsuarioNombre','FechaRevision','HoraRevision','TotalCitas','PendientesJSON','Observaciones','ResumenJSON','Actualizado'],
    RecordatoriosCita: ['ID','CitaID','FechaCita','Tipo','FechaEnvio','UsuarioID','UsuarioNombre','PendientePago','Mensaje','Canal','Estado'],
    NovedadesDiarias: ['ID','FechaObjetivo','CitaID','Tipo','Detalle','UsuarioID','UsuarioNombre','Fecha','Estado'],
    PlantillasPlanes: ['ID','Nombre','Descripcion','SesionesTotales','PrecioIndividual','PrecioTotal','PrecioSesionPlan','Descuento','NumeroCuotas','CuotasJSON','SesionesPorCuotaJSON','VigenciaDias','ServiciosIncluidos','Estado','Actualizado'],
    PlanesCliente: ['ID','Cliente','Telefono','Email','PlantillaID','NombrePlan','CitaOrigenID','SesionesTotales','SesionesPagadas','SesionesUsadas','SesionesDisponibles','SaldoPendiente','ProximaCuota','Vence','ProfesionalID','Estado','Creado','Actualizado'],
    CuotasPlan: ['ID','PlanClienteID','NumeroCuota','Valor','Estado','FechaPago','PagoID','SesionesHabilitadas','Vence','Observaciones'],
    SesionesPlan: ['ID','PlanClienteID','CitaID','NumeroSesion','Estado','ConsumeSesion','ProfesionalID','Fecha','Observaciones','Actualizado'],
    TarifasProfesionales: ['ID','ProfesionalID','Servicio','TipoTarifa','Valor','Porcentaje','Turno','Estado','Actualizado'],
    LiquidacionesProfesionales: ['ID','ProfesionalID','Periodo','Sesiones','Total','Estado','Creado','Pagado','Observaciones'],
    AuditoriaGeneral: ['ID','Fecha','UsuarioID','UsuarioNombre','Rol','Accion','Entidad','EntidadID','ValorAnterior','ValorNuevo','Motivo']
  };
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers[name].length).setValues([headers[name]]);
  } else if (headers[name]) {
    var lastCol = Math.max(sh.getLastColumn(), 1);
    var current = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return '' + (h || ''); });
    headers[name].forEach(function(h) {
      if (current.indexOf(h) === -1) {
        sh.getRange(1, sh.getLastColumn() + 1).setValue(h);
        current.push(h);
      }
    });
  }
  return sh;
}

function sheetObjects_(sh) {
  var values = sh.getDataRange().getValues();
  if (!values.length) return [];
  var headers = values[0], out = [];
  for (var i = 1; i < values.length; i++) {
    if (!values[i][0]) continue;
    var o = {};
    for (var j = 0; j < headers.length; j++) {
      var v = values[i][j];
      o[headers[j]] = v instanceof Date ? v.toISOString() : '' + (v || '');
    }
    out.push(o);
  }
  return out;
}

function auditGeneral_(user, action, entity, entityId, oldValue, newValue, reason) {
  user = user || {id:'system', nombre:'Sistema', rol:'Sistema'};
  operationsSheet_('AuditoriaGeneral').appendRow([
    'AUDG-' + new Date().getTime() + '-' + Math.floor(Math.random() * 999),
    new Date(), user.id || '', user.nombre || 'Sistema', user.rol || 'Sistema',
    action || '', entity || '', entityId || '',
    typeof oldValue === 'string' ? oldValue : JSON.stringify(oldValue || ''),
    typeof newValue === 'string' ? newValue : JSON.stringify(newValue || ''),
    reason || ''
  ]);
}

function upsertConfig_(key, value, description) {
  var sh = operationsSheet_('ConfiguracionOperativa'), rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if ('' + rows[i][0] === key) {
      sh.getRange(i + 1, 2, 1, 3).setValues([[value, description || rows[i][2] || '', new Date()]]);
      return;
    }
  }
  sh.appendRow([key, value, description || '', new Date()]);
}

function seedRole_(id, name, desc, permissions) {
  var sh = operationsSheet_('Roles'), rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) if ('' + rows[i][0] === id) return;
  sh.appendRow([id, name, desc, JSON.stringify(permissions || []), 'Activo', new Date(), new Date()]);
}

function seedPaymentAccount_(id, medio, tipo, numero, titular, order) {
  var sh = operationsSheet_('CuentasPago'), rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) if ('' + rows[i][0] === id) return;
  sh.appendRow([id, medio, tipo, numero, titular, 'Activa', order || 0, new Date()]);
}

function seedPlanTemplate_() {
  var sh = operationsSheet_('PlantillasPlanes'), rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) if ('' + rows[i][0] === 'PLAN-READAPTACION-6') return;
  sh.appendRow([
    'PLAN-READAPTACION-6','Plan de readaptaciÃ³n funcional','Plan base de 6 sesiones con pago en 2 cuotas.',
    6,70000,390000,65000,30000,2,
    JSON.stringify([{numero:1, valor:195000}, {numero:2, valor:195000}]),
    JSON.stringify([{cuota:1, sesiones:3}, {cuota:2, sesiones:3}]),
    60,'ReadaptaciÃ³n Funcional','Activo',new Date()
  ]);
}

function setupOperationsModule_() {
  [
    'Roles','UsuariosAdmin','CuentasPago','ConfiguracionOperativa','HistorialEstadosCita','Pagos','ComprobantesPago','SaldosFavor','Cortesias','AgendaPreparacion','RecordatoriosCita','NovedadesDiarias',
    'PlantillasPlanes','PlanesCliente','CuotasPlan','SesionesPlan','TarifasProfesionales','LiquidacionesProfesionales','AuditoriaGeneral'
  ].forEach(function(name) { operationsSheet_(name); });
  seedRole_('SUPERADMIN', 'Superadministradora', 'Acceso total del sistema', ['*']);
  seedRole_('ADMIN', 'Administrativa', 'Agenda, clientes, pagos, planes y reportes operativos', ['agenda:*','clientes:*','pagos:*','planes:*','reportes:operativos']);
  seedRole_('FISIO', 'Fisioterapeuta', 'Solo agenda propia y registro clÃ­nico sin finanzas', ['fisio:agenda','fisio:sesiones','fisio:novedades']);
  seedPaymentAccount_('CTA-BANCOLOMBIA', 'Bancolombia', 'Cuenta de ahorros', '91257857099', 'Jessica Andrea Ocampo Barbosa', 1);
  seedPaymentAccount_('CTA-NEQUI', 'Nequi', 'NÃºmero', '3136467945', 'Jessica Andrea Ocampo Barbosa', 2);
  seedPaymentAccount_('CTA-LLAVE', 'Llave', 'NÃºmero', '1010124692', 'Jessica Andrea Ocampo Barbosa', 3);
  upsertConfig_('reserva_temporal_minutos', '60', 'Tiempo inicial para mantener una reserva temporal pendiente de pago.');
  upsertConfig_('regla_atencion_confirmada', 'permitida', 'ExcepciÃ³n solicitada: una cita Confirmada puede atenderse cuando ya llegÃ³ la hora.');
  upsertConfig_('comprobantes_max_mb', '8', 'TamaÃ±o mÃ¡ximo sugerido para comprobantes JPG, JPEG, PNG o PDF.');
  seedPlanTemplate_();
  return {ok:true};
}

function reservationCodeFor_(citaId, fecha) {
  var f = fecha || fmtDate(new Date());
  var parts = f.split('-');
  var dmy = parts.length === 3 ? (parts[2] + parts[1] + ('' + parts[0]).slice(-2)) : Utilities.formatDate(new Date(), 'America/Bogota', 'ddMMyy');
  var raw = ('' + (citaId || '')).replace(/\D/g, '');
  var seq = raw ? raw.slice(-4) : ('' + Math.floor(Math.random() * 9999));
  while (seq.length < 4) seq = '0' + seq;
  return 'CF-' + dmy + '-' + seq;
}

function recordAppointmentStatusHistory_(citaId, prevState, nextState, user, obs) {
  var found = getCitaById_(citaId);
  var code = found ? reservationCodeFor_(citaId, found.cita.fecha) : reservationCodeFor_(citaId);
  user = user || {id:'admin', nombre:'Administracion', rol:'Superadministradora'};
  operationsSheet_('HistorialEstadosCita').appendRow([
    'HST-' + new Date().getTime() + '-' + Math.floor(Math.random()*999),
    citaId || '', code, prevState || '', nextState || '', new Date(),
    user.id || '', user.nombre || 'Administracion', user.rol || 'Superadministradora', obs || ''
  ]);
  auditGeneral_(user, 'Cambio estado cita', 'Cita', citaId, prevState || '', nextState || '', obs || '');
}

function recordReprogrammingHistory_(citaId, oldFecha, oldHora, newFecha, newHora, motivo, user, count) {
  var found = getCitaById_(citaId);
  var code = found ? reservationCodeFor_(citaId, found.cita.fecha) : reservationCodeFor_(citaId);
  user = user || {id:'admin', nombre:'Administracion', rol:'Superadministradora'};
  operationsSheet_('HistorialEstadosCita').appendRow([
    'HST-' + new Date().getTime() + '-' + Math.floor(Math.random()*999),
    citaId || '', code, 'REPROGRAMACION', 'REPROGRAMACION_REGISTRADA', new Date(),
    user.id || '', user.nombre || 'Administracion', user.rol || 'Superadministradora',
    motivo || '',
    'REPROGRAMACION',
    (oldFecha || '') + ' ' + (oldHora || ''),
    (newFecha || '') + ' ' + (newHora || ''),
    motivo || '',
    count || ''
  ]);
  auditGeneral_(user, 'Reprogramo cita', 'Cita', citaId, (oldFecha || '') + ' ' + (oldHora || ''), (newFecha || '') + ' ' + (newHora || ''), motivo || '');
}

function expireTemporaryReservations_() {
  setupOperationsModule_();
  var ss = getOrCreateSheet();
  var sheet = ss.getSheetByName('Citas');
  var map = ensureCitasStateColumns_();
  var rows = sheet.getDataRange().getValues();
  var now = new Date();
  var user = {id:'system', nombre:'Sistema', rol:'Automatizacion'};
  var expired = [];
  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    var citaId = '' + (row[0] || '');
    if (!citaId) continue;
    var state = getCitaStateFromRow_(row, map);
    var estadoPago = state.estadoPago || '';
    var estadoCita = state.estadoCita || '';
    var legacyState = '' + (row[10] || '');
    var hasLegacyPayment = hasLegacyPaymentMarker_(row[14]);
    var expiresAt = map.VenceReserva ? parseReservationExpiry_(row[map.VenceReserva - 1]) : null;
    if (!expiresAt) continue;
    if (expiresAt > now) continue;
    if (estadoPago !== PAYMENT_STATUS.PENDIENTE_PAGO) continue;
    if (estadoCita !== APPOINTMENT_STATUS.RESERVADA) continue;
    if (hasLegacyPayment) continue;
    if (hasActivePaymentForCita_(citaId)) continue;
    var prev = legacyState + ' => ' + estadoPago + ' / ' + estadoCita;
    var next = 'RESERVA_VENCIDA => ' + estadoPago + ' / ' + APPOINTMENT_STATUS.CANCELADA;
    sheet.getRange(i + 1, 11).setValue('Cancelada');
    if (map.EstadoCita) sheet.getRange(i + 1, map.EstadoCita).setValue(APPOINTMENT_STATUS.CANCELADA);
    if (map.RequiereCierreAdmin) sheet.getRange(i + 1, map.RequiereCierreAdmin).setValue('');
    recordAppointmentStatusHistory_(citaId, prev, next, user, 'Motivo: RESERVA_VENCIDA. Reserva temporal vencida sin comprobante ni pago aprobado.');
    auditGeneral_(user, 'RESERVA_VENCIDA', 'Cita', citaId, prev, next, 'Cancelada por vencimiento de 60 minutos; horario liberado.');
    expired.push({id:citaId, nombre:row[2] || '', fecha:sd(row[7]), hora:st(row[8])});
  }
  return {ok:true, expired:expired.length, citas:expired};
}

function hasLegacyPaymentMarker_(value) {
  var raw = ('' + (value || '')).trim();
  if (!raw) return false;
  var normalizedStatuses = [
    PAYMENT_STATUS.PENDIENTE_PAGO,
    PAYMENT_STATUS.COMPROBANTE_RECIBIDO,
    PAYMENT_STATUS.PAGO_APROBADO,
    PAYMENT_STATUS.PAGO_RECHAZADO,
    PAYMENT_STATUS.REEMBOLSADO,
    PAYMENT_STATUS.NO_REQUIERE_PAGO
  ];
  if (normalizedStatuses.indexOf(raw) >= 0) return false;
  var key = normalizeKey_(raw);
  if (key === 'pendiente' || key === 'pendiente de pago') return false;
  return true;
}

function parseReservationExpiry_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  if (typeof value === 'number' && value > 0) {
    var epoch = new Date(Math.round((value - 25569) * 86400 * 1000));
    if (!isNaN(epoch.getTime())) return epoch;
  }
  var raw = ('' + (value || '')).trim();
  if (!raw) return null;
  var parsed = new Date(raw);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function hasActivePaymentForCita_(citaId) {
  setupOperationsModule_();
  var rows = operationsSheet_('Pagos').getDataRange().getValues();
  var active = [
    PAYMENT_STATUS.COMPROBANTE_RECIBIDO,
    PAYMENT_STATUS.PAGO_APROBADO,
    PAYMENT_STATUS.NO_REQUIERE_PAGO,
    'Por verificar',
    'Aprobado'
  ];
  for (var i = 1; i < rows.length; i++) {
    if ('' + (rows[i][2] || '') !== '' + citaId) continue;
    if (active.indexOf('' + (rows[i][12] || '')) >= 0) return true;
  }
  var cort = operationsSheet_('Cortesias').getDataRange().getValues();
  for (var j = 1; j < cort.length; j++) {
    if ('' + (cort[j][1] || '') === '' + citaId && normalizeKey_(cort[j][8]) === 'autorizada') return true;
  }
  return false;
}

function doUpdateStatus(p, user) {
  setupOperationsModule_();
  ensureCitasStateColumns_();
  var sheet = getOrCreateSheet().getSheetByName('Citas');
  var map = headerMap_(sheet);
  var rows  = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === p.id) {
      var prev = '' + (rows[i][10] || '');
      sheet.getRange(i+1, 11).setValue(p.status);
      var statusKey = normalizeKey_(p.status);
      var state = getCitaStateFromRow_(rows[i], map);
      var nextPago = state.estadoPago;
      var nextCita = state.estadoCita;
      if (statusKey === 'atendida' || statusKey === 'sesion atendida') nextCita = APPOINTMENT_STATUS.ATENDIDA;
      else if (statusKey === 'no asistio') nextCita = APPOINTMENT_STATUS.NO_ASISTIO;
      else if (statusKey === 'cancelada' || statusKey.indexOf('cancel') === 0) nextCita = APPOINTMENT_STATUS.CANCELADA;
      else if (statusKey === 'pago rechazado') nextPago = PAYMENT_STATUS.PAGO_RECHAZADO;
      else if (statusKey === 'pago por verificar') nextPago = PAYMENT_STATUS.COMPROBANTE_RECIBIDO;
      else if (statusKey === 'pago verificado' || statusKey === 'autorizada para atender') {
        nextPago = PAYMENT_STATUS.PAGO_APROBADO;
        nextCita = APPOINTMENT_STATUS.AUTORIZADA;
      }
      if (map.EstadoPago) sheet.getRange(i+1, map.EstadoPago).setValue(nextPago);
      if (map.EstadoCita) sheet.getRange(i+1, map.EstadoCita).setValue(nextCita);
      if (nextCita === APPOINTMENT_STATUS.ATENDIDA && !p.manualAudit) {
        if (p.note) sheet.getRange(i+1, 14).setValue(p.note);
      }
      if (p.note) sheet.getRange(i+1, 14).setValue(p.note);
      var nextRow = rows[i].slice();
      if (map.EstadoPago) nextRow[map.EstadoPago - 1] = nextPago;
      if (map.EstadoCita) nextRow[map.EstadoCita - 1] = nextCita;
      try { passportSyncAppointment_(nextRow, map, user || {id:'admin', nombre:'Administracion', rol:'Superadministradora'}); } catch(passportErr) {
        try { auditGeneral_(user, 'Error sincronizando Pasaporte', 'Cita', p.id, '', passportErr.message, p.note || ''); } catch(auditErr) {}
      }
      recordAppointmentStatusHistory_(p.id, prev, p.status + ' => ' + nextPago + ' / ' + nextCita, user || {id:'admin', nombre:'Administracion', rol:'Superadministradora'}, p.note || 'Cambio desde agenda admin');
      return {ok: true};
    }
  }
  return {ok: false, error: 'Cita no encontrada'};
}

function parseOperationsPayload_(data) {
  if (!data) return {};
  if (typeof data === 'object') return data;
  try {
    return JSON.parse(decodeURIComponent(data));
  } catch(e1) {
    try { return JSON.parse(data); } catch(e2) { return {}; }
  }
}

function operationConfigValue_(key, fallback) {
  var rows = operationsSheet_('ConfiguracionOperativa').getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if ('' + rows[i][0] === key) return rows[i][1] || fallback;
  }
  return fallback;
}

function paymentProofFolder_() {
  var name = 'Comprobantes Cuidandote Fisioterapia';
  var folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(name);
}

function hexDigest_(bytes) {
  return bytes.map(function(b) {
    var v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? '0' + v : v;
  }).join('');
}

function savePaymentProof_(file, meta, user) {
  if (!file || !file.data) return null;
  var allowed = {'image/jpeg': true, 'image/png': true, 'application/pdf': true};
  var mime = '' + (file.type || '');
  var name = ('' + (file.name || 'comprobante')).replace(/[\\\/:*?"<>|]/g, '-');
  if (!allowed[mime]) return {ok:false,error:'El comprobante debe ser JPG, PNG o PDF'};

  var raw = '' + file.data;
  var comma = raw.indexOf(',');
  if (comma > -1) raw = raw.slice(comma + 1);
  var bytes;
  try {
    bytes = Utilities.base64Decode(raw);
  } catch(e) {
    return {ok:false,error:'No se pudo leer el comprobante. Intenta subirlo de nuevo.'};
  }
  var maxMb = Number(operationConfigValue_('comprobantes_max_mb', '8')) || 8;
  if (bytes.length > maxMb * 1024 * 1024) return {ok:false,error:'El comprobante supera el tamaÃ±o mÃ¡ximo de ' + maxMb + ' MB'};

  var digest = hexDigest_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, bytes));
  var proofSh = operationsSheet_('ComprobantesPago');
  var proofRows = proofSh.getDataRange().getValues();
  for (var i = 1; i < proofRows.length; i++) {
    var sameHash = ('' + (proofRows[i][11] || '')) === digest;
    var sameCita = ('' + (proofRows[i][3] || '')) === ('' + (meta.citaId || ''));
    if (sameHash && sameCita) return {ok:false,error:'Este comprobante ya fue registrado para esta cita'};
  }

  var driveFile = paymentProofFolder_().createFile(Utilities.newBlob(bytes, mime, name)).setName((meta.codigoReserva || meta.pagoId || 'comprobante') + ' - ' + name);
  try { driveFile.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE); } catch(e) {}

  var proofId = 'PRF-' + new Date().getTime() + '-' + Math.floor(Math.random() * 999);
  proofSh.appendRow([
    proofId, meta.pagoId || '', meta.codigoReserva || '', meta.citaId || '',
    name, mime, bytes.length, driveFile.getId(), 'Recibido', new Date(),
    meta.observaciones || '', digest
  ]);
  auditGeneral_(user, 'Cargo comprobante de pago', 'ComprobantePago', proofId, '', {pagoId:meta.pagoId, citaId:meta.citaId, archivo:name}, '');
  return {ok:true,id:proofId,fileId:driveFile.getId(),url:driveFile.getUrl(),hash:digest};
}

function upsertProfessionalSettlement_(professionalId, citaId, servicio, tarifa, attendedAt, user) {
  setupOperationsModule_();
  if (!professionalId || !citaId) return;
  var period = Utilities.formatDate(attendedAt || new Date(), 'America/Bogota', 'yyyy-MM');
  var value = Number(('' + (tarifa || '')).replace(/[^\d.-]/g, '')) || 0;
  var sh = operationsSheet_('LiquidacionesProfesionales'), rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if ('' + rows[i][1] === '' + professionalId && '' + rows[i][2] === period && ['Pendiente de liquidacion','Pendiente de liquidaciÃ³n',''].indexOf('' + (rows[i][5] || '')) > -1) {
      var sessions = Number(rows[i][3] || 0) + 1;
      var total = Number(rows[i][4] || 0) + value;
      var obs = (rows[i][8] ? rows[i][8] + '\n' : '') + citaId + ' Â· ' + (servicio || '') + ' Â· ' + value;
      sh.getRange(i + 1, 4, 1, 6).setValues([[sessions, total, 'Pendiente de liquidacion', rows[i][6] || new Date(), rows[i][7] || '', obs]]);
      auditGeneral_(user, 'Actualizo liquidacion profesional', 'LiquidacionProfesional', rows[i][0], '', {periodo:period, sesiones:sessions, total:total}, citaId);
      return;
    }
  }
  var id = 'LIQ-' + new Date().getTime() + '-' + Math.floor(Math.random() * 999);
  sh.appendRow([id, professionalId, period, 1, value, 'Pendiente de liquidacion', new Date(), '', citaId + ' Â· ' + (servicio || '') + ' Â· ' + value]);
  auditGeneral_(user, 'Creo liquidacion profesional', 'LiquidacionProfesional', id, '', {periodo:period, sesiones:1, total:value}, citaId);
}

function savePayment_(data, user) {
  setupOperationsModule_();
  ensureCitasStateColumns_();
  var p = parseOperationsPayload_(data);
  if (!p.citaId && !p.codigoReserva) return {ok:false,error:'Falta cita o cÃ³digo de reserva'};
  var found = p.citaId ? getCitaById_(p.citaId) : null;
  var code = p.codigoReserva || (found ? reservationCodeFor_(p.citaId, found.cita.fecha) : reservationCodeFor_(''));
  var id = p.id || ('PAY-' + new Date().getTime());
  var proofResult = p.proofFile ? savePaymentProof_(p.proofFile, {
    pagoId: id,
    codigoReserva: code,
    citaId: p.citaId || '',
    observaciones: p.observaciones || ''
  }, user) : null;
  if (proofResult && !proofResult.ok) return proofResult;
  if (proofResult && proofResult.url && !p.comprobante) p.comprobante = proofResult.url;
  var sh = operationsSheet_('Pagos'), rows = sh.getDataRange().getValues(), row = -1;
  for (var i = 1; i < rows.length; i++) if ('' + rows[i][0] === id) row = i + 1;
  if (row < 0 && p.citaId) {
    for (var j = rows.length - 1; j >= 1; j--) {
      var sameCita = '' + (rows[j][2] || '') === '' + p.citaId;
      var activeStatus = '' + (rows[j][12] || '');
      if (sameCita && ['COMPROBANTE_RECIBIDO','PENDIENTE_PAGO','Por verificar'].indexOf(activeStatus) >= 0) {
        row = j + 1;
        id = '' + rows[j][0];
        break;
      }
    }
  }
  var expected = p.valorEsperado || (found ? found.cita.precio : '');
  var values = [
    id, code, p.citaId || '', p.cliente || (found ? found.cita.nombre : ''),
    p.servicioPlan || (found ? found.cita.servicio : ''), expected, p.valorRecibido || '',
    p.medioPago || '', p.cuentaReceptora || '', p.fechaPago || '', p.fechaVerificacion || '',
    p.comprobante || '', p.estadoPago || PAYMENT_STATUS.COMPROBANTE_RECIBIDO, p.usuarioVerifico || '',
    p.observaciones || '', p.cuotaNumero || '', p.saldoPendiente || '', row > 0 ? rows[row-1][17] || new Date() : new Date(), new Date()
  ];
  if (row > 0) sh.getRange(row, 1, 1, values.length).setValues([values]);
  else sh.appendRow(values);
  if (p.citaId && (p.estadoPago || PAYMENT_STATUS.COMPROBANTE_RECIBIDO) === PAYMENT_STATUS.COMPROBANTE_RECIBIDO) {
    var f = getCitaById_(p.citaId);
    if (f) {
      getOrCreateSheet().getSheetByName('Citas').getRange(f.row, 11).setValue('Pago por verificar');
      setCitaStates_(p.citaId, PAYMENT_STATUS.COMPROBANTE_RECIBIDO, APPOINTMENT_STATUS.RESERVADA, user, 'Comprobante registrado pendiente de verificacion');
    }
  }
  auditGeneral_(user, row > 0 ? 'ActualizÃ³ pago' : 'RegistrÃ³ pago', 'Pago', id, '', values, p.observaciones || '');
  return {ok:true,id:id,codigoReserva:code};
}

function verifyPayment_(p, user) {
  setupOperationsModule_();
  ensureCitasStateColumns_();
  var id = p.id, status = p.estado || p.status || '';
  if (!id || !status) return {ok:false,error:'Falta pago o estado'};
  var statusKey = normalizeKey_(status);
  var normalizedPay = status;
  if (statusKey === 'aprobado' || statusKey === 'pago aprobado') normalizedPay = PAYMENT_STATUS.PAGO_APROBADO;
  else if (statusKey === 'rechazado' || statusKey === 'pago rechazado') normalizedPay = PAYMENT_STATUS.PAGO_RECHAZADO;
  else if (statusKey === 'por verificar' || statusKey === 'comprobante recibido') normalizedPay = PAYMENT_STATUS.COMPROBANTE_RECIBIDO;
  else if (statusKey === 'reembolsado' || statusKey === 'reembolsada') normalizedPay = PAYMENT_STATUS.REEMBOLSADO;
  var sh = operationsSheet_('Pagos'), rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if ('' + rows[i][0] !== '' + id) continue;
    var prevPay = '' + (rows[i][12] || '');
    sh.getRange(i+1, 11, 1, 5).setValues([[new Date(), rows[i][11] || '', normalizedPay, user.nombre || 'Administracion', p.observaciones || rows[i][14] || '']]);
    sh.getRange(i+1, 19).setValue(new Date());
    var citaId = '' + (rows[i][2] || '');
    if (citaId) {
      var found = getCitaById_(citaId);
      if (found) {
        var nextState = normalizedPay === PAYMENT_STATUS.PAGO_APROBADO ? 'Autorizada para atender' : (normalizedPay === PAYMENT_STATUS.PAGO_RECHAZADO ? 'Pago rechazado' : (normalizedPay === PAYMENT_STATUS.COMPROBANTE_RECIBIDO ? 'Pago por verificar' : found.cita.estado));
        getOrCreateSheet().getSheetByName('Citas').getRange(found.row, 11).setValue(nextState);
        if (normalizedPay === PAYMENT_STATUS.PAGO_APROBADO) getOrCreateSheet().getSheetByName('Citas').getRange(found.row, 15).setValue(rows[i][7] || 'Pago aprobado');
        var currentState = getCitaStateFromRow_(found.raw || [], headerMap_(getOrCreateSheet().getSheetByName('Citas')));
        var nextCitaStatus = currentState.estadoCita || APPOINTMENT_STATUS.RESERVADA;
        if (normalizedPay === PAYMENT_STATUS.PAGO_APROBADO) nextCitaStatus = APPOINTMENT_STATUS.AUTORIZADA;
        if (normalizedPay === PAYMENT_STATUS.PAGO_RECHAZADO) nextCitaStatus = APPOINTMENT_STATUS.RESERVADA;
        if (normalizedPay === PAYMENT_STATUS.REEMBOLSADO) nextCitaStatus = APPOINTMENT_STATUS.CANCELADA;
        setCitaStates_(citaId, normalizedPay, nextCitaStatus, user, 'Verificacion de pago: ' + normalizedPay);
      }
    }
    auditGeneral_(user, 'Verifico pago', 'Pago', id, prevPay, normalizedPay, p.observaciones || '');
    return {ok:true};
  }
  return {ok:false,error:'Pago no encontrado'};
}

function authorizeCourtesy_(p, user) {
  setupOperationsModule_();
  ensureCitasStateColumns_();
  if (!p.citaId) return {ok:false,error:'Falta cita'};
  if (!p.motivo) return {ok:false,error:'Debes registrar el motivo de la cortesÃ­a'};
  var found = getCitaById_(p.citaId);
  if (!found) return {ok:false,error:'Cita no encontrada'};
  var valor = p.valorComercial || found.cita.precio || '';
  if (!valor) return {ok:false,error:'Debes registrar el valor comercial del servicio'};
  var sh = operationsSheet_('Cortesias');
  var id = 'COR-' + new Date().getTime() + '-' + Math.floor(Math.random()*999);
  sh.appendRow([id, p.citaId, found.cita.nombre, found.cita.servicio, valor, p.motivo, user.nombre || 'Administracion', new Date(), 'Autorizada', p.observaciones || '']);
  getOrCreateSheet().getSheetByName('Citas').getRange(found.row, 11).setValue('CortesÃ­a autorizada');
  setCitaStates_(p.citaId, PAYMENT_STATUS.NO_REQUIERE_PAGO, APPOINTMENT_STATUS.AUTORIZADA, user, 'Cortesia autorizada: ' + p.motivo + ' | Valor comercial: ' + valor);
  auditGeneral_(user, 'Autorizo cortesia', 'Cortesia', id, '', {citaId:p.citaId, valorComercial:valor, motivo:p.motivo}, p.observaciones || '');
  return {ok:true,id:id};
}

function createCreditBalance_(p, user) {
  setupOperationsModule_();
  var valor = Number(('' + (p.valorDisponible || p.valor || '')).replace(/[^\d.-]/g,'')) || 0;
  if (!valor) return {ok:false,error:'Debes registrar el valor disponible'};
  var id = p.id || ('SALDO-' + new Date().getTime() + '-' + Math.floor(Math.random()*999));
  operationsSheet_('SaldosFavor').appendRow([
    id, p.cliente || '', p.telefono || '', valor, p.citaOrigenId || '', p.pagoOrigenId || '',
    new Date(), p.fechaVencimiento || '', 0, valor, user.nombre || 'Administracion',
    'Disponible', p.observaciones || '', new Date()
  ]);
  auditGeneral_(user, 'Creo saldo a favor', 'SaldoFavor', id, '', {valorDisponible:valor, citaOrigen:p.citaOrigenId || '', pagoOrigen:p.pagoOrigenId || ''}, p.observaciones || '');
  return {ok:true,id:id,saldoRestante:valor};
}

function applyCreditBalance_(p, user) {
  setupOperationsModule_();
  ensureCitasStateColumns_();
  if (!p.saldoId || !p.citaId) return {ok:false,error:'Falta saldo o cita'};
  var sh = operationsSheet_('SaldosFavor'), rows = sh.getDataRange().getValues();
  var valorUsar = Number(('' + (p.valorUtilizado || p.valor || '')).replace(/[^\d.-]/g,'')) || 0;
  for (var i = 1; i < rows.length; i++) {
    if ('' + rows[i][0] !== '' + p.saldoId) continue;
    var restante = Number(rows[i][9] || 0);
    if (!valorUsar) valorUsar = restante;
    if (valorUsar > restante) return {ok:false,error:'El valor a usar supera el saldo disponible'};
    var nuevoRestante = restante - valorUsar;
    var usado = Number(rows[i][8] || 0) + valorUsar;
    sh.getRange(i+1, 9, 1, 5).setValues([[usado, nuevoRestante, rows[i][10] || user.nombre || 'Administracion', nuevoRestante > 0 ? 'Disponible' : 'Usado', rows[i][12] || '']]);
    sh.getRange(i+1, 14).setValue(new Date());
    var payResult = savePayment_({
      citaId:p.citaId,
      valorRecibido:valorUsar,
      medioPago:'SALDO_A_FAVOR',
      cuentaReceptora:p.saldoId,
      fechaPago:fmtDate(new Date()),
      estadoPago:PAYMENT_STATUS.PAGO_APROBADO,
      observaciones:'Pago aplicado con saldo a favor ' + p.saldoId
    }, user);
    if (payResult && payResult.id) verifyPayment_({id:payResult.id, estado:PAYMENT_STATUS.PAGO_APROBADO}, user);
    setCitaStates_(p.citaId, PAYMENT_STATUS.PAGO_APROBADO, APPOINTMENT_STATUS.AUTORIZADA, user, 'Saldo a favor aplicado: ' + p.saldoId + ' valor ' + valorUsar);
    auditGeneral_(user, 'Aplico saldo a favor', 'SaldoFavor', p.saldoId, restante, nuevoRestante, 'Cita ' + p.citaId);
    return {ok:true,saldoRestante:nuevoRestante,valorUtilizado:valorUsar};
  }
  return {ok:false,error:'Saldo a favor no encontrado'};
}

function savePaymentAccount_(data, user) {
  setupOperationsModule_();
  var a = JSON.parse(decodeURIComponent(data || '{}'));
  if (!a.medio || !a.numero) return {ok:false,error:'Falta medio o nÃºmero'};
  var id = a.id || ('CTA-' + new Date().getTime());
  var sh = operationsSheet_('CuentasPago'), rows = sh.getDataRange().getValues(), row = -1;
  for (var i = 1; i < rows.length; i++) if ('' + rows[i][0] === id) row = i + 1;
  var values = [id, a.medio, a.tipo || '', a.numero, a.titular || 'Jessica Andrea Ocampo Barbosa', a.estado || 'Activa', a.orden || 9, new Date()];
  if (row > 0) sh.getRange(row, 1, 1, values.length).setValues([values]);
  else sh.appendRow(values);
  auditGeneral_(user, row > 0 ? 'ActualizÃ³ cuenta de pago' : 'CreÃ³ cuenta de pago', 'CuentaPago', id, '', values, '');
  return {ok:true,id:id};
}

function getOperationsData_() {
  setupOperationsModule_();
  return {
    ok:true,
    estados: APPOINTMENT_STATUS_CATALOG,
    cuentas: sheetObjects_(operationsSheet_('CuentasPago')),
    config: sheetObjects_(operationsSheet_('ConfiguracionOperativa')),
    pagos: sheetObjects_(operationsSheet_('Pagos')).reverse(),
    saldosFavor: sheetObjects_(operationsSheet_('SaldosFavor')).reverse(),
    cortesias: sheetObjects_(operationsSheet_('Cortesias')).reverse(),
    agendaPreparacion: sheetObjects_(operationsSheet_('AgendaPreparacion')).reverse(),
    recordatoriosCita: sheetObjects_(operationsSheet_('RecordatoriosCita')).reverse(),
    novedadesDiarias: sheetObjects_(operationsSheet_('NovedadesDiarias')).reverse(),
    historialEstados: sheetObjects_(operationsSheet_('HistorialEstadosCita')).slice(-120).reverse(),
    plantillasPlanes: sheetObjects_(operationsSheet_('PlantillasPlanes')),
    planesCliente: sheetObjects_(operationsSheet_('PlanesCliente')),
    cuotasPlan: sheetObjects_(operationsSheet_('CuotasPlan')),
    sesionesPlan: sheetObjects_(operationsSheet_('SesionesPlan')),
    tarifasProfesionales: sheetObjects_(operationsSheet_('TarifasProfesionales')),
    liquidaciones: sheetObjects_(operationsSheet_('LiquidacionesProfesionales')),
    auditoria: sheetObjects_(operationsSheet_('AuditoriaGeneral')).slice(-120).reverse()
  };
}

function parseISODate_(s) {
  if (!s) return null;
  if (Object.prototype.toString.call(s) === '[object Date]') return s;
  var parts = ('' + s).split('-').map(Number);
  if (parts.length < 3) return null;
  return new Date(parts[0], parts[1] - 1, parts[2]);
}

function addDaysISO_(dateStr, days) {
  var d = parseISODate_(dateStr) || new Date();
  d.setDate(d.getDate() + days);
  return Utilities.formatDate(d, 'America/Bogota', 'yyyy-MM-dd');
}

function normalizedPaymentStateForDaily_(cita) {
  var state = cita.estadoPago || cita.EstadoPago || '';
  if (state) return '' + state;
  return legacyPaymentToNew_(cita.estado || cita.Estado || '', cita.pago || cita.Pago || '').status;
}

function normalizedAppointmentStateForDaily_(cita, pago, assigned) {
  var state = cita.estadoCita || cita.EstadoCita || '';
  if (state) return '' + state;
  return legacyAppointmentToNew_(cita.estado || cita.Estado || '', pago, assigned).status;
}

function buildDailyOpsSummary_(targetDate, mode) {
  var assignments = getAssignmentMap_();
  var reminders = {};
  sheetObjects_(operationsSheet_('RecordatoriosCita')).forEach(function(r) {
    var id = '' + (r.CitaID || r.citaId || '');
    if (id) reminders[id] = r;
  });
  var preparations = sheetObjects_(operationsSheet_('AgendaPreparacion')).reverse().filter(function(p) {
    return sd(p.FechaObjetivo || p.fechaObjetivo || '') === '' + targetDate;
  });
  var lastPrep = preparations[0] || null;
  var prepTime = lastPrep ? new Date(lastPrep.Actualizado || lastPrep.FechaRevision || lastPrep.fechaRevision || 0) : null;
  var historial = [];
  var todayNow = new Date();
  var sh = getOrCreateSheet().getSheetByName('Citas');
  var rows = sh ? sh.getDataRange().getValues() : [];
  var head = rows.length ? rows[0].map(function(h){ return '' + (h || ''); }) : [];
  function idx(name, fallback) { var i = head.indexOf(name); return i >= 0 ? i : fallback; }
  var ix = {
    id:idx('ID',0), nombre:idx('Nombre',2), telefono:idx('Telefono',3), email:idx('Email',4),
    servicio:idx('Servicio',5), modalidad:idx('Modalidad',6), fecha:idx('FechaCita',7), hora:idx('Hora',8),
    precio:idx('Precio',9), estado:idx('Estado',10), direccion:idx('Direccion',11), notas:idx('Notas',12),
    notaAdmin:idx('NotaAdmin',13), pago:idx('Pago',14), estadoPago:idx('EstadoPago',-1), estadoCita:idx('EstadoCita',-1),
    reprogramaciones:idx('Reprogramaciones',-1), planClienteId:idx('PlanClienteID',-1)
  };
  var targetCitas = [];
  for (var r = 1; r < rows.length; r++) {
    var row = rows[r];
    if (!row[ix.id] || sd(row[ix.fecha]) !== targetDate) continue;
    targetCitas.push({
      id:row[ix.id], nombre:row[ix.nombre], telefono:row[ix.telefono], email:row[ix.email],
      servicio:row[ix.servicio], modalidad:row[ix.modalidad], fecha:sd(row[ix.fecha]), hora:st(row[ix.hora]),
      precio:row[ix.precio], estado:row[ix.estado], direccion:row[ix.direccion], notas:row[ix.notas],
      notaAdmin:row[ix.notaAdmin], pago:row[ix.pago],
      estadoPago:ix.estadoPago >= 0 ? row[ix.estadoPago] : '',
      estadoCita:ix.estadoCita >= 0 ? row[ix.estadoCita] : '',
      Reprogramaciones:ix.reprogramaciones >= 0 ? row[ix.reprogramaciones] : 0,
      PlanClienteID:ix.planClienteId >= 0 ? row[ix.planClienteId] : ''
    });
  }
  var nowMinutes = todayNow.getHours() * 60 + todayNow.getMinutes();
  var inThreeHours = nowMinutes + 180;
  if (mode === 'hoy') {
    targetCitas.sort(function(a,b){ return st(a.hora || a.Hora).localeCompare(st(b.hora || b.Hora)); });
  }
  var items = targetCitas.map(function(c) {
    var id = '' + (c.id || c.ID || '');
    var assigned = assignments[id] || {};
    var pago = normalizedPaymentStateForDaily_(c);
    var cita = normalizedAppointmentStateForDaily_(c, pago, !!assigned.ProfesionalID || !!assigned.profesionalId);
    var reminder = reminders[id] || null;
    var h = st(c.hora || c.Hora);
    var minutes = Number(h.split(':')[0] || 0) * 60 + Number(h.split(':')[1] || 0);
    var changedAfterPrep = false;
    if (prepTime) {
      changedAfterPrep = historial.some(function(ev) {
        return '' + (ev.CitaID || ev.citaId || '') === id && new Date(ev.Fecha || ev.fecha || 0) > prepTime;
      });
    }
    var pendingClose = false;
    if (mode === 'hoy' && minutes < nowMinutes && [APPOINTMENT_STATUS.ATENDIDA, APPOINTMENT_STATUS.NO_ASISTIO, APPOINTMENT_STATUS.CANCELADA].indexOf(cita) === -1) pendingClose = true;
    return {
      id:id,
      hora:h,
      nombre:c.nombre || c.Nombre || '',
      telefono:c.telefono || c.Telefono || '',
      servicio:c.servicio || c.Servicio || '',
      modalidad:c.modalidad || c.Modalidad || '',
      sede:c.sede || c.Sede || c.modalidad || c.Modalidad || '',
      valor:c.precio || c.Precio || '',
      estadoPago:pago,
      estadoCita:cita,
      confirmadaVisual:isPaymentAuthorizing_(pago) && [APPOINTMENT_STATUS.AUTORIZADA, APPOINTMENT_STATUS.ASIGNADA].indexOf(cita) >= 0,
      fisioterapeuta:assigned.ProfesionalNombre || assigned.profesionalNombre || '',
      fisioterapeutaId:assigned.ProfesionalID || assigned.profesionalId || '',
      observaciones:c.notaAdmin || c.NotaAdmin || c.notas || c.Notas || '',
      reprogramaciones:c.Reprogramaciones || c.reprogramaciones || 0,
      plan:!!(c.PlanClienteID || c.planClienteId || (c.servicio || '').toLowerCase().indexOf('plan') >= 0 || (c.servicio || '').toLowerCase().indexOf('paquete') >= 0),
      recordatorio: reminder ? 'ENVIADO' : 'PENDIENTE',
      proximaTresHoras: mode === 'hoy' && minutes >= nowMinutes && minutes <= inThreeHours,
      pendienteCierre: pendingClose,
      cambioPosteriorPreparacion: changedAfterPrep
    };
  });
  var counts = {
    total: items.length,
    autorizadas: items.filter(function(i){ return isPaymentAuthorizing_(i.estadoPago) && [APPOINTMENT_STATUS.AUTORIZADA, APPOINTMENT_STATUS.ASIGNADA].indexOf(i.estadoCita) >= 0; }).length,
    pendientesPago: items.filter(function(i){ return i.estadoPago === PAYMENT_STATUS.PENDIENTE_PAGO; }).length,
    comprobantes: items.filter(function(i){ return i.estadoPago === PAYMENT_STATUS.COMPROBANTE_RECIBIDO; }).length,
    sinFisioterapeuta: items.filter(function(i){ return !i.fisioterapeuta && [APPOINTMENT_STATUS.CANCELADA, APPOINTMENT_STATUS.NO_ASISTIO].indexOf(i.estadoCita) === -1; }).length,
    recordatoriosPendientes: items.filter(function(i){ return i.recordatorio !== 'ENVIADO'; }).length,
    novedades: items.filter(function(i){ return i.cambioPosteriorPreparacion || i.pendienteCierre; }).length,
    pendientesCierre: items.filter(function(i){ return i.pendienteCierre; }).length,
    cambiosPosteriores: items.filter(function(i){ return i.cambioPosteriorPreparacion; }).length,
    listasAtender: items.filter(function(i){ return i.confirmadaVisual && i.fisioterapeuta; }).length,
    datosIncompletos: items.filter(function(i){ return !i.telefono || !i.servicio || !i.hora; }).length
  };
  return {
    ok:true,
    mode:mode || 'manana',
    fecha:targetDate,
    preparacion:lastPrep,
    counts:counts,
    citas:items,
    horariosDisponibles:[],
    generatedAt:new Date()
  };
}

function getDailyOperationsData_(date, mode) {
  var target = date || (mode === 'hoy' ? Utilities.formatDate(new Date(), 'America/Bogota', 'yyyy-MM-dd') : addDaysISO_(Utilities.formatDate(new Date(), 'America/Bogota', 'yyyy-MM-dd'), 1));
  return buildDailyOpsSummary_(target, mode || 'manana');
}

function saveDailyPreparation_(data, user) {
  setupOperationsModule_();
  var d = JSON.parse(decodeURIComponent(data || '{}'));
  if (!d.fechaObjetivo) return {ok:false,error:'Falta fecha objetivo'};
  var summary = buildDailyOpsSummary_(d.fechaObjetivo, d.tipo || 'manana');
  var pendientes = {
    pendientesPago: summary.counts.pendientesPago,
    comprobantes: summary.counts.comprobantes,
    sinFisioterapeuta: summary.counts.sinFisioterapeuta,
    recordatoriosPendientes: summary.counts.recordatoriosPendientes,
    datosIncompletos: summary.counts.datosIncompletos,
    conflictosHorario: 0
  };
  var estado = Object.keys(pendientes).some(function(k){ return Number(pendientes[k]) > 0; }) ? 'CON_PENDIENTES' : 'LISTA';
  var id = d.id || ('PREP-' + new Date().getTime());
  operationsSheet_('AgendaPreparacion').appendRow([
    id, d.fechaObjetivo, d.tipo || 'manana', d.estadoPreparacion || estado, user.id || '', user.nombre || '',
    Utilities.formatDate(new Date(), 'America/Bogota', 'yyyy-MM-dd'), Utilities.formatDate(new Date(), 'America/Bogota', 'HH:mm:ss'),
    summary.counts.total, JSON.stringify(pendientes), d.observaciones || '', JSON.stringify(summary.counts), new Date()
  ]);
  auditGeneral_(user, 'MarcÃ³ agenda como revisada', 'AgendaPreparacion', id, '', estado, d.observaciones || '');
  return {ok:true,id:id,estado:estado,pendientes:pendientes,summary:summary.counts};
}

function logDailyReminder_(p, user) {
  setupOperationsModule_();
  if (!p.citaId) return {ok:false,error:'Falta cita'};
  var id = 'REC-' + new Date().getTime();
  operationsSheet_('RecordatoriosCita').appendRow([
    id, p.citaId, p.fecha || '', p.tipo || 'recordatorio_cita', new Date(), user.id || '', user.nombre || '',
    p.pendientePago === 'true' || p.pendientePago === true, p.mensaje || '', 'WhatsApp', 'REGISTRADO'
  ]);
  auditGeneral_(user, 'RegistrÃ³ recordatorio de cita', 'RecordatorioCita', id, '', p.citaId, '');
  return {ok:true,id:id};
}

function addDailyObservation_(p, user) {
  setupOperationsModule_();
  var id = 'NOV-' + new Date().getTime();
  operationsSheet_('NovedadesDiarias').appendRow([
    id, p.fecha || '', p.citaId || '', p.tipo || 'OBSERVACION', p.detalle || '', user.id || '', user.nombre || '', new Date(), 'ABIERTA'
  ]);
  auditGeneral_(user, 'AgregÃ³ novedad diaria', 'NovedadDiaria', id, '', p.detalle || '', '');
  return {ok:true,id:id};
}

function infoPlan(serv, mod) {
  var s = (serv || '').split('(')[0].trim();
  var esDom = mod === 'Domicilio';
  var planes = {
    'Paquete ReadaptaciÃ³n Inicio': { total:6,  pagoDosEn:4, mitadP:'$189.000', mitadD:'$234.500' },
    'Paquete ReadaptaciÃ³n Avance': { total:8,  pagoDosEn:5, mitadP:'$238.000', mitadD:'$299.000' },
    'Paquete ReadaptaciÃ³n Total':  { total:10, pagoDosEn:6, mitadP:'$280.000', mitadD:'$361.000' },
    'Paquete RecuperaciÃ³n Full':   { total:3,  pagoDosEn:null, mitadP:null, mitadD:null },
    'Combo DiagnÃ³stico Pro':       { total:2,  pagoDosEn:null, mitadP:null, mitadD:null },
    'Combo Bienvenida':            { total:2,  pagoDosEn:null, mitadP:null, mitadD:null },
    'Plan Activo':                 { total:2,  pagoDosEn:null, mitadP:null, mitadD:null },
    'Plan Pro':                    { total:3,  pagoDosEn:null, mitadP:null, mitadD:null },
    'Mini-sesiÃ³n Familiar 20 min': { total:1,  pagoDosEn:null, mitadP:null, mitadD:null },
  };
  for (var k in planes) {
    if (s === k || s.indexOf(k) === 0) {
      var p = planes[k];
      return { total: p.total, pagoDosEn: p.pagoDosEn, mitad: esDom ? p.mitadD : p.mitadP };
    }
  }
  var sl = s.toLowerCase();
  if (sl.indexOf('paquete') === 0 || sl.indexOf('plan ') === 0 || sl.indexOf('combo') === 0 || sl.indexOf('mini') === 0) {
    return { total: null, pagoDosEn: null, mitad: null };
  }
  return null;
}

function contarSesiones(rows, nombre, serv, excludeFecha) {
  var norm    = (nombre || '').toLowerCase().trim();
  var planKey = (serv || '').split('(')[0].trim();
  var count   = 0;
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (r[10] === 'Cancelada') continue;
    var rNombre  = ('' + (r[2] || '')).toLowerCase().trim();
    var rPlanKey = ('' + (r[5] || '')).split('(')[0].trim();
    var rFecha   = (r[7] instanceof Date) ? fmtDate(r[7]) : ('' + (r[7]||'')).split('T')[0];
    if (rNombre === norm && rPlanKey === planKey && rFecha < excludeFecha) count++;
  }
  return count;
}

function mensajePlanWA(nombre, serv, hora, mod, fechaLegible, plan, sesionActual, esHoy) {
  var primerNombre = nombre.split(' ')[0];
  var cuando = esHoy ? 'Hoy ' + fechaLegible : 'MaÃ±ana ' + fechaLegible;
  var encabezado = esHoy ? 'ðŸ©º *Â¡Hoy tienes cita!*' : 'ðŸ©º *Recordatorio de cita*';
  var progLine = (plan && sesionActual)
    ? '\nðŸ”„ SesiÃ³n ' + sesionActual + (plan.total ? ' de ' + plan.total : '')
    : '';

  var pagoLine = '';
  if (esHoy) {
    if (plan && sesionActual) {
      if (sesionActual === 1) {
        pagoLine = '\n\nðŸ’³ Pago inicial' + (plan.mitad ? ': ' + plan.mitad : '') +
          '\nBancolombia Ahorros: 91257857099\nLlave: 1010124692\nNequi: 3136467945\nTitular: Jessica Andrea Ocampo Barbosa';
      } else if (plan.pagoDosEn && sesionActual === plan.pagoDosEn) {
        pagoLine = '\n\nðŸ’³ Segundo pago del plan' + (plan.mitad ? ': ' + plan.mitad : '') +
          '\nBancolombia Ahorros: 91257857099\nLlave: 1010124692\nNequi: 3136467945\nTitular: Jessica Andrea Ocampo Barbosa';
      }
    }
  }

  return encabezado + '\nâ”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”\n' +
    'ðŸ’† ' + serv + '\n' +
    'ðŸ“… ' + cuando + '\n' +
    'ðŸ•˜ ' + hora + ' Â· ' + mod +
    progLine +
    pagoLine + '\n\n' +
    'Â¿Tienes algÃºn cambio? EscrÃ­beme ðŸ™\n' +
    '_Cuidándote Fisioterapia_';
}

// -------------------------------------------------------------
//  RECORDATORIOS DIARIOS â€” ejecutar con trigger 7am
// -------------------------------------------------------------
function sendReminders() {
  var diasSemana = ['domingo','lunes','martes','miÃ©rcoles','jueves','viernes','sÃ¡bado'];
  var meses = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];

  var ss   = getOrCreateSheet();
  var rows = ss.getSheetByName('Citas').getDataRange().getValues();

  var today    = fmtDate(new Date());
  var tmrwDate = new Date(); tmrwDate.setDate(tmrwDate.getDate() + 1);
  var tomorrow = fmtDate(tmrwDate);

  var linksHoy = [], linksManana = [];

  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (r[10] === 'Cancelada' || r[10] === 'Atendida') continue;
    var fecha = (r[7] instanceof Date) ? fmtDate(r[7]) : ('' + (r[7]||'')).split('T')[0];
    var nombre = r[2], email = r[4];
    var serv  = r[5], mod = r[6], precio = r[9];
    var hora  = st(r[8]);
    var rawTel = (r[3] instanceof Error) ? '' : ('' + (r[3]||''));
    var phone  = rawTel.replace(/\D/g,'');
    if (phone.length <= 10) phone = '57' + phone;

    // Detectar si es plan y calcular sesiÃ³n actual
    var plan = infoPlan(serv, mod);
    var sesionActual = null;
    if (plan) {
      sesionActual = contarSesiones(rows, nombre, serv, fecha) + 1;
    }

    if (fecha === tomorrow) {
      var dp = fecha.split('-');
      var fObj = new Date(+dp[0], +dp[1]-1, +dp[2]);
      var fechaLegible = diasSemana[fObj.getDay()] + ' ' + +dp[2] + ' de ' + meses[+dp[1]-1];
      if (email && email.indexOf('@') > 0) {
        GmailApp.sendEmail(
          email,
          'Recordatorio: maÃ±ana tienes cita â€” Cuidándote Fisioterapia',
          'Este correo requiere un cliente de correo con soporte HTML.',
          {htmlBody: buildReminderEmail(nombre, serv, fechaLegible, hora, mod, precio, false, plan, sesionActual),
           name: 'Cuidándote Fisioterapia'}
        );
      }
      var msg1 = mensajePlanWA(nombre, serv, hora, mod, fechaLegible, plan, sesionActual, false);
      linksManana.push(nombre + ' (' + hora + '): https://wa.me/' + phone + '?text=' + encodeURIComponent(msg1));
      try { queueAutomationMessage_('Recordatorio', nombre, phone, msg1, 'Cita maÃ±ana ' + hora, r[0], 'appt-tomorrow|' + r[0] + '|' + today); } catch(q1) {}
    }

    if (fecha === today) {
      var dp2 = fecha.split('-');
      var fObj2 = new Date(+dp2[0], +dp2[1]-1, +dp2[2]);
      var fechaLegible2 = diasSemana[fObj2.getDay()] + ' ' + +dp2[2] + ' de ' + meses[+dp2[1]-1];
      if (email && email.indexOf('@') > 0) {
        GmailApp.sendEmail(
          email,
          'â° Hoy tienes cita â€” Cuidándote Fisioterapia',
          'Este correo requiere un cliente de correo con soporte HTML.',
          {htmlBody: buildReminderEmail(nombre, serv, fechaLegible2, hora, mod, precio, true, plan, sesionActual),
           name: 'Cuidándote Fisioterapia'}
        );
      }
      var msg2 = mensajePlanWA(nombre, serv, hora, mod, fechaLegible2, plan, sesionActual, true);
      linksHoy.push(nombre + ' (' + hora + '): https://wa.me/' + phone + '?text=' + encodeURIComponent(msg2));
      try { queueAutomationMessage_('Recordatorio', nombre, phone, msg2, 'Cita hoy ' + hora, r[0], 'appt-today|' + r[0] + '|' + today); } catch(q2) {}
    }
  }

  // Resumen diario para Jessica con links de WhatsApp 1-clic
  if (linksHoy.length > 0 || linksManana.length > 0) {
    var body = 'Recordatorios automÃ¡ticos del dÃ­a ' + today + '\n\n';
    if (linksHoy.length)    body += '== CITAS DE HOY (WhatsApp 1 clic) ==\n' + linksHoy.join('\n') + '\n\n';
    if (linksManana.length) body += '== CITAS DE MANANA (WhatsApp 1 clic) ==\n' + linksManana.join('\n') + '\n';
    GmailApp.sendEmail(JESSICA_EMAIL, 'Resumen de citas - ' + today, body);
  }
}

// Ejecuta ESTA funcion UNA sola vez para activar los recordatorios diarios:
function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('sendReminders').timeBased().everyDays(1).atHour(7).inTimezone('America/Bogota').create();
  ScriptApp.newTrigger('expireTemporaryReservations_').timeBased().everyMinutes(15).create();
  Logger.log('Triggers activados: sendReminders 7am hora Colombia. autoMarcarAtendidas permanece desactivado.');
}

function setupReservationExpirationTrigger_() {
  var exists = false;
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'expireTemporaryReservations_') exists = true;
  });
  if (!exists) ScriptApp.newTrigger('expireTemporaryReservations_').timeBased().everyMinutes(15).create();
  return {ok:true, active:true, handler:'expireTemporaryReservations_'};
}

// -------------------------------------------------------------
//  AUTO MARCAR ATENDIDAS â€” trigger diario a las 10pm
// -------------------------------------------------------------
function autoMarcarAtendidas() {
  return {
    ok: true,
    disabled: true,
    count: 0,
    message: 'autoMarcarAtendidas desactivado: las citas solo se cierran por fisioterapeuta o administradora con auditoria.'
  };
  var ss    = getOrCreateSheet();
  var sheet = ss.getSheetByName('Citas');
  var rows  = sheet.getDataRange().getValues();
  var ahora = new Date();
  var hoy   = fmtDate(ahora);
  var count = 0;

  for (var i = 1; i < rows.length; i++) {
    var r      = rows[i];
    var estado = ('' + (r[10] || '')).trim();
    // Solo aplica a citas Confirmadas o Pendientes â€” no tocar Canceladas, Atendidas, No asistiÃ³
    if (estado !== 'Confirmada' && estado !== 'Pendiente') continue;

    var fecha = (r[7] instanceof Date) ? fmtDate(r[7]) : ('' + r[7]).split('T')[0];
    var hora  = st(r[8]);

    // Citas de dÃ­as anteriores â†’ marcar como Atendida directamente
    if (fecha < hoy) {
      sheet.getRange(i + 1, 11).setValue('Atendida');
      count++;
      continue;
    }

    // Citas de hoy â†’ marcar solo si la hora ya pasÃ³ (+ 30 min de margen)
    if (fecha === hoy) {
      var parts    = hora.split(':');
      var citaFin  = new Date();
      citaFin.setHours(parseInt(parts[0], 10), parseInt(parts[1], 10) + 30, 0, 0);
      if (ahora > citaFin) {
        sheet.getRange(i + 1, 11).setValue('Atendida');
        count++;
      }
    }
  }

  Logger.log('autoMarcarAtendidas: ' + count + ' citas marcadas como Atendida.');
  return { ok: true, count: count };
}

// -------------------------------------------------------------
//  HELPERS
// -------------------------------------------------------------
function getOrCreateSheet() {
  var files = DriveApp.getFilesByName(SS_NAME);
  if (files.hasNext()) {
    var ss = SpreadsheetApp.open(files.next());
    // Crear hoja Pacientes si no existe aÃºn
    if (!ss.getSheetByName('Pacientes')) {
      ss.insertSheet('Pacientes').getRange(1,1,1,5).setValues([[
        'Nombre','Telefono','Email','PrimeraVisita','UltimaVisita'
      ]]);
    }
    return ss;
  }

  var ss = SpreadsheetApp.create(SS_NAME);
  var cs = ss.getActiveSheet(); cs.setName('Citas');
  cs.getRange(1,1,1,15).setValues([[
    'ID','FechaRegistro','Nombre','Telefono','Email',
    'Servicio','Modalidad','FechaCita','Hora','Precio',
    'Estado','Direccion','Notas','NotaAdmin','Pago'
  ]]);
  ss.insertSheet('Bloqueos').getRange(1,1,1,6).setValues([[
    'Fecha','HoraInicio','HoraFin','Motivo','CreadoPor','ID'
  ]]);
  ss.insertSheet('Pacientes').getRange(1,1,1,5).setValues([[
    'Nombre','Telefono','Email','PrimeraVisita','UltimaVisita'
  ]]);
  return ss;
}

function upsertPaciente(nombre, telefono, email) {
  try {
    var ss    = getOrCreateSheet();
    var sheet = ss.getSheetByName('Pacientes');
    var phone = ('' + (telefono || '')).replace(/\D/g, '');
    var norm  = (nombre || '').toLowerCase().trim();
    var today = new Date().toLocaleDateString('es-CO');
    var data  = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      var rowNorm  = ('' + (data[i][0] || '')).toLowerCase().trim();
      var rowPhone = ('' + (data[i][1] || '')).replace(/\D/g, '');
      if (rowNorm === norm || (phone && rowPhone === phone)) {
        // Actualizar telÃ©fono/email si llegaron nuevos y actualizar Ãºltima visita
        if (phone && !rowPhone)    sheet.getRange(i+1, 2).setValue(phone);
        if (email && !data[i][2])  sheet.getRange(i+1, 3).setValue(email);
        sheet.getRange(i+1, 5).setValue(today);
        return;
      }
    }
    // Nuevo paciente
    sheet.appendRow([nombre, phone, email || '', today, today]);
    sheet.getRange(sheet.getLastRow(), 2).setNumberFormat('@').setValue(phone);
  } catch(e) {
    Logger.log('upsertPaciente error: ' + e.message);
    throw e;
  }
}

function getServiceDuration(service) {
  // Servicios combinados: "Descarga + ReadaptaciÃ³n" â†’ suma de duraciones
  if (service && service.indexOf(' + ') !== -1) {
    return service.split(' + ').reduce(function(sum, s) { return sum + getServiceDuration(s.trim()); }, 0);
  }
  var s = (service || '').toLowerCase()
    .replace(/[Ã¡Ã Ã¢]/g,'a').replace(/[Ã©Ã¨Ãª]/g,'e')
    .replace(/[Ã­Ã¬Ã®]/g,'i').replace(/[Ã³Ã²Ã´]/g,'o').replace(/[ÃºÃ¹Ã»]/g,'u');
  if (s.indexOf('mini') > -1)           return 20;  // Mini-sesiÃ³n Familiar 20 min
  if (s.indexOf('completa') > -1)       return 90;  // Descarga Muscular Completa: visible 1h20, bloqueo 90 min
  if (s.indexOf('full') > -1)           return 80;  // Paquete RecuperaciÃ³n Full
  if (s.indexOf('plan pro') > -1)       return 80;  // Plan Pro (sesiÃ³n Full incluida)
  if (s.indexOf('paquete readap') > -1) return 45;  // Paquetes ReadaptaciÃ³n Inicio/Avance/Total
  if (s.indexOf('readaptacion') > -1)   return 60;  // Readaptación Funcional: visible 40-45, bloqueo 60 min
  if (s.indexOf('valoracion') > -1)     return 60;  // Valoración Funcional: visible 50, bloqueo 60 min
  if (s.indexOf('piernas') > -1)        return 60;  // Descarga Muscular Piernas: visible 50, bloqueo 60 min
  if (s.indexOf('cuello') > -1)         return 60;  // Descarga Muscular Cuello/Espalda: visible 50, bloqueo 60 min
  if (s.indexOf('plan activo') > -1)    return 50;  // Plan Activo (Express)
  if (s.indexOf('combo') > -1)          return 50;  // Combo DiagnÃ³stico Pro / Combo Bienvenida
  return 60;
}

function parseDT(date, time) {
  var d = date.split('-'), t = time.split(':');
  return new Date(+d[0], +d[1]-1, +d[2], +t[0], +t[1]);
}

function fmtDate(d) {
  return d.getFullYear() + '-' + pad(d.getMonth()+1) + '-' + pad(d.getDate());
}

function pad(n) { return n < 10 ? '0'+n : ''+n; }

function txt(s) { return ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.TEXT); }
function js(o)  { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function buildDesc(d, price) {
  return 'PACIENTE: ' + d.name + '\nTelefono: ' + d.phone + '\nEmail: ' + d.email +
    '\n---\nServicio: ' + d.service + '\nModalidad: ' + d.modality +
    '\nDireccion: ' + (d.address || 'Presencial') + '\nValor: ' + price +
    '\n---\nNotas: ' + (d.notes || 'Sin notas') + '\n\nAgendado desde jessicaocampoft-ctrl.github.io';
}

function buildEmailJessica(d, price) {
  return 'Nueva cita agendada desde tu pagina web!\n\n' +
    'SERVICIO: ' + d.service + '\nFecha: ' + d.date + ' a las ' + d.time +
    '\nModalidad: ' + d.modality + '\nValor: ' + price +
    '\n\nPACIENTE\nNombre: ' + d.name + '\nTelefono: ' + d.phone + '\nEmail: ' + d.email +
    (d.address ? '\nDireccion: ' + d.address : '') +
    (d.notes   ? '\nNotas: '     + d.notes   : '');
}

function buildEmailCliente(d, price) {
  var diasSemana = ['domingo','lunes','martes','miÃ©rcoles','jueves','viernes','sÃ¡bado'];
  var meses = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
  var dp = d.date.split('-');
  var fechaObj = new Date(+dp[0], +dp[1]-1, +dp[2]);
  var fechaLegible = diasSemana[fechaObj.getDay()] + ' ' + +dp[2] + ' de ' + meses[+dp[1]-1] + ' de ' + dp[0];
  var modDetalle = d.modality + (d.address ? ' â€” ' + d.address : '');
  var primerNombre = d.name.split(' ')[0];
  var codigoReserva = d.codigoReserva || '';

  return '<div style="font-family:Arial,sans-serif;max-width:520px;margin:0 auto;border:1px solid #e5e7eb;border-radius:12px;overflow:hidden">' +
    '<div style="background:#0d9488;padding:20px 32px;text-align:center">' +
    '<p style="color:#fff;margin:0;font-size:15px;font-weight:600">ðŸ©º Cuidándote Fisioterapia</p>' +
    '</div>' +
    '<div style="padding:28px 32px">' +
    '<p style="font-size:17px;font-weight:700;margin:0 0 12px;color:#111827">Reserva temporal creada, ' + primerNombre + '</p>' +
    '<p style="font-size:13px;color:#6b7280;margin:0 0 18px;line-height:1.6">Tu horario queda reservado por 60 minutos. Para confirmar la cita debes realizar el pago anticipado y enviar el comprobante para verificacion administrativa.</p>' +
    '<div style="margin:0 0 20px">' +
    '<p style="margin:0 0 6px;font-size:14px;font-weight:600;color:#111827">ðŸ“Œ ' + d.service + '</p>' +
    '<p style="margin:0;font-size:13px;color:#6b7280">' + fechaLegible + ' Â· ' + d.time + ' Â· ' + modDetalle + '</p>' +
    '</div>' +
    '<hr style="border:none;border-top:2px solid #e5e7eb;margin:20px 0">' +
    '<div style="font-size:13px;color:#374151;line-height:1.7;margin:0 0 16px">' +
    '<p style="margin:0 0 8px;font-weight:700;color:#111827">Datos de pago</p>' +
    '<p style="margin:0">Codigo de reserva: ' + codigoReserva + '</p>' +
    '<p style="margin:0">Valor: ' + price + '</p>' +
    '<p style="margin:8px 0 0">Bancolombia - Cuenta de ahorros: 91257857099</p>' +
    '<p style="margin:0">Nequi: 3136467945</p>' +
    '<p style="margin:0">Llave: 1010124692</p>' +
    '<p style="margin:8px 0 0">Titular: Jessica Andrea Ocampo Barbosa</p>' +
    '<p style="margin:8px 0 0;color:#6b7280">Despues de pagar, envia el comprobante junto con tu nombre completo y codigo de reserva.</p>' +
    '</div>' +
    '<hr style="border:none;border-top:1px solid #e5e7eb;margin:20px 0">' +
    '<p style="font-size:14px;color:#374151;margin:0 0 4px">Hasta pronto. Gracias por confiar en nuestros servicios.</p>' +
    '<p style="font-size:14px;color:#374151;margin:0;font-style:italic">â€” Cuidándote Fisioterapia</p>' +
    '</div>' +
    '<div style="background:#f9fafb;padding:14px 32px;text-align:center;font-size:12px;color:#9ca3af">' +
    'Cuidándote Fisioterapia Â· Pereira, Colombia Â· ' +
    '<a href="https://wa.me/573136467945" style="color:#0d9488">+57 313 646 7945</a>' +
    '</div></div>';
}

function buildPaymentBlock(titulo, subtitulo) {
  return '<div style="background:#fffbeb;border:1px solid #fde68a;border-radius:8px;padding:14px 18px;margin:0 0 4px">' +
    '<p style="margin:0 0 ' + (subtitulo ? '6px' : '10px') + ';font-size:14px;font-weight:700;color:#92400e">' + titulo + '</p>' +
    (subtitulo ? '<p style="margin:0 0 10px;font-size:13px;color:#78350f">' + subtitulo + '</p>' : '') +
    '<table style="width:100%;font-size:13px;color:#44403c;border-collapse:collapse">' +
    '<tr><td style="padding:3px 0;color:#78716c;width:140px">Bancolombia Ahorros</td><td style="padding:3px 0;font-weight:600">91257857099</td></tr>' +
    '<tr><td style="padding:3px 0;color:#78716c">Llave</td><td style="padding:3px 0;font-weight:600">1010124692</td></tr>' +
    '<tr><td style="padding:3px 0;color:#78716c">Nequi</td><td style="padding:3px 0;font-weight:600">3136467945</td></tr>' +
    '<tr><td style="padding:3px 0;color:#78716c">Titular</td><td style="padding:3px 0">Jessica Andrea Ocampo Barbosa</td></tr>' +
    '</table></div>';
}

function buildReminderEmail(nombre, serv, fechaLegible, hora, mod, precio, esHoy, plan, sesionActual) {
  var primerNombre = nombre.split(' ')[0];

  var intro = esHoy
    ? 'Â¡Hola ' + primerNombre + '! ðŸŒŸ Hoy es el dÃ­a de tu sesiÃ³n de <strong>' + serv + '</strong> a las <strong>' + hora + '</strong> (' + mod + ').'
    : 'Â¡Hola ' + primerNombre + '! ðŸ‘‹ Te recuerdo que maÃ±ana tienes tu sesiÃ³n de <strong>' + serv + '</strong> a las <strong>' + hora + '</strong> (' + mod + ').';

  var progStr = '';
  if (plan && sesionActual) {
    progStr = '<p style="margin:10px 0 0;font-size:13px;color:#6b7280">ðŸ”„ ' +
      (plan.total ? 'SesiÃ³n ' + sesionActual + ' de ' + plan.total : 'SesiÃ³n ' + sesionActual) + '</p>';
  }

  // Bloque de pago: solo en recordatorio del mismo dÃ­a
  var bloquePago = '';
  if (esHoy) {
    if (plan && sesionActual) {
      if (sesionActual === 1) {
        bloquePago = '<hr style="border:none;border-top:1px solid #e5e7eb;margin:20px 0">' +
          buildPaymentBlock('ðŸ’° Pago inicial del plan' + (plan.mitad ? ' â€” ' + plan.mitad : ''), 'Para comenzar recuerda traer el pago inicial.');
      } else if (plan.pagoDosEn && sesionActual === plan.pagoDosEn) {
        bloquePago = '<hr style="border:none;border-top:1px solid #e5e7eb;margin:20px 0">' +
          buildPaymentBlock('ðŸ’° Segundo pago del plan' + (plan.mitad ? ' â€” ' + plan.mitad : ''), 'Esta sesiÃ³n corresponde al segundo y Ãºltimo pago de tu plan. Recuerda traerlo.');
      }
    } else if (!plan && precio) {
      bloquePago = '<hr style="border:none;border-top:1px solid #e5e7eb;margin:20px 0">' +
        buildPaymentBlock('ðŸ’° Valor: ' + precio, null);
    }
  }

  return '<div style="font-family:Arial,sans-serif;max-width:520px;margin:0 auto;border:1px solid #e5e7eb;border-radius:12px;overflow:hidden">' +
    '<div style="background:' + (esHoy ? '#0284c7' : '#0d9488') + ';padding:20px 32px;text-align:center">' +
    '<p style="color:#fff;margin:0;font-size:15px;font-weight:600">ðŸ©º Cuidándote Fisioterapia</p>' +
    '</div>' +
    '<div style="padding:28px 32px">' +
    '<p style="font-size:14px;line-height:1.7;color:#111827;margin:0">' + intro + '</p>' +
    progStr +
    bloquePago +
    '<hr style="border:none;border-top:1px solid #e5e7eb;margin:20px 0">' +
    '<p style="font-size:13px;color:#6b7280;margin:0 0 16px">Â¿Tienes algÃºn cambio? <a href="https://wa.me/573136467945" style="color:#0d9488">EscrÃ­beme</a>.</p>' +
    '<p style="font-size:14px;color:#374151;margin:0 0 4px">Hasta pronto. Gracias por confiar en nuestros servicios.</p>' +
    '<p style="font-size:14px;color:#374151;margin:0;font-style:italic">â€” Cuidándote Fisioterapia</p>' +
    '</div>' +
    '<div style="background:#f9fafb;padding:14px 32px;text-align:center;font-size:12px;color:#9ca3af">' +
    'Cuidándote Fisioterapia Â· Pereira, Colombia Â· ' +
    '<a href="https://wa.me/573136467945" style="color:#0d9488">+57 313 646 7945</a>' +
    '</div></div>';
}

// =============================================================
//  ADMIN KV â€” almacenamiento clave-valor sincronizado entre dispositivos
// =============================================================

function getAdminKVSheet() {
  var ss = getOrCreateSheet();
  var sh = ss.getSheetByName('AdminKV');
  if (!sh) {
    sh = ss.insertSheet('AdminKV');
    sh.appendRow(['key', 'value', 'updated']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function getAdminKV() {
  try {
    var sh   = getAdminKVSheet();
    var data = sh.getDataRange().getValues();
    var kv   = {};
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) kv['' + data[i][0]] = '' + data[i][1];
    }
    return { ok: true, kv: kv };
  } catch(e) {
    return { ok: false, error: e.message, kv: {} };
  }
}

function doSetAdminKV(dataJson) {
  try {
    var updates = JSON.parse(decodeURIComponent(dataJson));
    var sh   = getAdminKVSheet();
    var data = sh.getDataRange().getValues();
    var now  = new Date().toISOString();

    // Construir Ã­ndice key â†’ nÃºmero de fila (1-based)
    var keyToRow = {};
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) keyToRow['' + data[i][0]] = i + 1;
    }

    for (var key in updates) {
      var val = updates[key];
      if (val === '__DELETE__') {
        if (keyToRow[key]) {
          sh.deleteRow(keyToRow[key]);
          // Reconstruir Ã­ndice tras borrar
          data = sh.getDataRange().getValues();
          keyToRow = {};
          for (var j = 1; j < data.length; j++) {
            if (data[j][0]) keyToRow['' + data[j][0]] = j + 1;
          }
        }
      } else if (keyToRow[key]) {
        sh.getRange(keyToRow[key], 2, 1, 2).setValues([[val, now]]);
      } else {
        sh.appendRow([key, val, now]);
        keyToRow[key] = sh.getLastRow();
      }
    }
    return { ok: true };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// =============================================================
//  PASAPORTE DE MOVIMIENTO
// =============================================================

function getPasaportesSheet() {
  var ss = getOrCreateSheet();
  var sh = ss.getSheetByName('Pasaportes');
  if (!sh) {
    sh = ss.insertSheet('Pasaportes');
    sh.appendRow(['nombre', 'passport', 'descarga', 'actualizado', 'PasaporteID', 'TokenPublico', 'Estado', 'FechaCreacionToken', 'FechaUltimoAcceso', 'Telefono', 'CitasAplicadasJSON']);
    sh.setFrozenRows(1);
  }
  ensurePasaportesSchema_(sh);
  return sh;
}

function ensurePasaportesSchema_(sh) {
  var required = ['nombre', 'passport', 'descarga', 'actualizado', 'PasaporteID', 'TokenPublico', 'Estado', 'FechaCreacionToken', 'FechaUltimoAcceso', 'Telefono', 'CitasAplicadasJSON'];
  var headers = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(function(h){ return '' + (h || ''); });
  required.forEach(function(h) {
    if (headers.indexOf(h) === -1) {
      sh.getRange(1, sh.getLastColumn() + 1).setValue(h);
      headers.push(h);
    }
  });
}

function passportHeaderMap_(sh) {
  ensurePasaportesSchema_(sh);
  return headerMap_(sh);
}

function makePassportId_() {
  return 'PAS-' + Utilities.formatDate(new Date(), 'America/Bogota', 'yyyyMMdd') + '-' + Utilities.getUuid().slice(0, 8).toUpperCase();
}

function makePassportToken_() {
  return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 12);
}

function passportPublicUrl_(id, token) {
  return 'https://cuidandotefisioterapia.com/pasaporte.html?id=' + encodeURIComponent(id || '') + '&token=' + encodeURIComponent(token || '');
}

function parseJsonSafe_(raw) {
  try { return raw ? JSON.parse(raw) : {}; } catch(e) { return {}; }
}

function passportAppliedAppointments_(row, map) {
  return map.CitasAplicadasJSON ? parseJsonSafe_(row[map.CitasAplicadasJSON - 1]) : {};
}

function passportActiveAppointmentCount_(applied) {
  return Object.keys(applied || {}).filter(function(id) { return applied[id] && applied[id].activo; }).length;
}

function passportLegacyManualCount_(passport) {
  var stamps = passport && (passport.stamps || passport.sellos);
  if (Array.isArray(stamps)) return stamps.filter(Boolean).length;
  if (stamps && typeof stamps === 'object') return Object.keys(stamps).filter(function(k){ return !!stamps[k]; }).length;
  return 0;
}

function passportEffectiveStampCount_(passport, applied) {
  var manual = passport && typeof passport.manualAdjustment === 'number'
    ? passport.manualAdjustment
    : passportLegacyManualCount_(passport);
  return Math.max(0, Math.min(16, passportActiveAppointmentCount_(applied) + manual));
}

function passportStampsForCount_(count) {
  var stamps = {};
  for (var i = 1; i <= 16; i++) stamps[i] = i <= count;
  return stamps;
}

function passportProgressForClient_(passport, applied) {
  var safe = passport || {};
  var effective = passportEffectiveStampCount_(safe, applied);
  return {
    stamps: passportStampsForCount_(effective),
    autoStampCount: passportActiveAppointmentCount_(applied),
    manualAdjustment: typeof safe.manualAdjustment === 'number' ? safe.manualAdjustment : passportLegacyManualCount_(safe)
  };
}

function passportRowToObject_(row, map, rowNumber) {
  var id = '' + (row[map.PasaporteID - 1] || '');
  var token = '' + (row[map.TokenPublico - 1] || '');
  return {
    row: rowNumber,
    nombre: '' + (row[map.nombre - 1] || ''),
    telefono: map.Telefono ? '' + (row[map.Telefono - 1] || '') : '',
    passport: passportProgressForClient_(parseJsonSafe_(row[map.passport - 1]), passportAppliedAppointments_(row, map)),
    descarga: parseJsonSafe_(row[map.descarga - 1]),
    actualizado: row[map.actualizado - 1] || '',
    id: id,
    token: token,
    estado: '' + (row[map.Estado - 1] || 'ACTIVO'),
    fechaCreacionToken: row[map.FechaCreacionToken - 1] || '',
    fechaUltimoAcceso: row[map.FechaUltimoAcceso - 1] || '',
    link: passportPublicUrl_(id, token)
  };
}

function migratePasaportesIfNeeded_() {
  var sh = getPasaportesSheet();
  var map = passportHeaderMap_(sh);
  var rows = sh.getDataRange().getValues();
  var changed = 0;
  for (var i = 1; i < rows.length; i++) {
    var now = new Date();
    if (!rows[i][map.PasaporteID - 1]) {
      sh.getRange(i + 1, map.PasaporteID).setValue(makePassportId_());
      changed++;
    }
    if (!rows[i][map.TokenPublico - 1]) {
      sh.getRange(i + 1, map.TokenPublico).setValue(makePassportToken_());
      sh.getRange(i + 1, map.FechaCreacionToken).setValue(now);
      changed++;
    }
    if (!rows[i][map.Estado - 1]) {
      sh.getRange(i + 1, map.Estado).setValue('ACTIVO');
      changed++;
    }
  }
  return {ok:true, migrated:changed};
}

function getLegacyPassportNotice_() {
  return {
    ok: false,
    legacy: true,
    error: 'Tu enlace de Pasaporte Cuidándote fue actualizado. Solicita tu nuevo enlace por WhatsApp.',
    whatsapp: 'https://wa.me/573136467945?text=' + encodeURIComponent('Hola, necesito mi nuevo enlace de Pasaporte Cuidándote Fisioterapia.')
  };
}

function getPassportSecure_(id, token) {
  try {
    if (!id || !token) return {ok:false,error:'Enlace incompleto. Solicita tu nuevo enlace por WhatsApp.'};
    var sh   = getPasaportesSheet();
    var map  = passportHeaderMap_(sh);
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if ('' + (data[i][map.PasaporteID - 1] || '') === '' + id) {
        if ('' + (data[i][map.TokenPublico - 1] || '') !== '' + token) return {ok:false,error:'Enlace inválido o actualizado. Solicita tu nuevo enlace por WhatsApp.'};
        if (('' + (data[i][map.Estado - 1] || 'ACTIVO')).toUpperCase() !== 'ACTIVO') return {ok:false,error:'Este pasaporte no está activo. Escríbenos por WhatsApp.'};
        sh.getRange(i + 1, map.FechaUltimoAcceso).setValue(new Date());
        var rawPassport = parseJsonSafe_(data[i][map.passport - 1]);
        var applied = passportAppliedAppointments_(data[i], map);
        return {
          ok:       true,
          id:       id,
          nombre:   '' + (data[i][map.nombre - 1] || ''),
          passport: passportProgressForClient_(rawPassport, applied),
          descarga: parseJsonSafe_(data[i][map.descarga - 1])
        };
      }
    }
    return { ok: false, error: 'Pasaporte no encontrado. Solicita tu nuevo enlace por WhatsApp.' };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

function findPassportByName_(nombre) {
  var sh = getPasaportesSheet();
  var map = passportHeaderMap_(sh);
  var rows = sh.getDataRange().getValues();
  var norm = (nombre || '').toLowerCase().trim();
  for (var i = 1; i < rows.length; i++) {
    if (('' + (rows[i][map.nombre - 1] || '')).toLowerCase().trim() === norm) {
      return passportRowToObject_(rows[i], map, i + 1);
    }
  }
  return null;
}

function findPassportById_(id) {
  var sh = getPasaportesSheet();
  var map = passportHeaderMap_(sh);
  var rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if ('' + (rows[i][map.PasaporteID - 1] || '') === '' + id) return passportRowToObject_(rows[i], map, i + 1);
  }
  return null;
}

function passportAppointmentIsEligible_(citaRow, citaMap) {
  var state = getCitaStateFromRow_(citaRow, citaMap);
  var service = '' + (citaRow[5] || '');
  var price = Number(citaRow[9] || 0);
  return state.estadoCita === APPOINTMENT_STATUS.ATENDIDA &&
    state.estadoPago === PAYMENT_STATUS.PAGO_APROBADO &&
    price > 0 && !esRegistro(service);
}

function passportSyncAppointment_(citaRow, citaMap, user) {
  var citaId = '' + (citaRow[0] || '');
  var nombre = '' + (citaRow[2] || '');
  if (!citaId || !nombre) return {ok:false, skipped:true};
  var eligible = passportAppointmentIsEligible_(citaRow, citaMap);
  var passport = findPassportByName_(nombre);
  if (!passport && !eligible) return {ok:true, skipped:true};
  if (!passport) {
    var created = passportEnsure_(nombre, ('' + (citaRow[3] || '')).replace(/\D/g, ''), user);
    if (!created.ok) return created;
    passport = findPassportById_(created.passport.id);
  }
  var sh = getPasaportesSheet();
  var map = passportHeaderMap_(sh);
  var row = sh.getRange(passport.row, 1, 1, sh.getLastColumn()).getValues()[0];
  var applied = passportAppliedAppointments_(row, map);
  var oldCount = passportEffectiveStampCount_(parseJsonSafe_(row[map.passport - 1]), applied);
  var current = applied[citaId] || {};
  if (!!current.activo === eligible) return {ok:true, changed:false};
  applied[citaId] = {
    citaId:citaId,
    servicio:'' + (citaRow[5] || ''),
    fecha:citaRow[7] || '',
    activo:eligible,
    actualizado:new Date(),
    usuarioId:(user && user.id) || 'system',
    usuarioNombre:(user && user.nombre) || 'Sistema'
  };
  sh.getRange(passport.row, map.CitasAplicadasJSON).setValue(JSON.stringify(applied));
  sh.getRange(passport.row, map.actualizado).setValue(new Date());
  var newCount = passportEffectiveStampCount_(parseJsonSafe_(row[map.passport - 1]), applied);
  try {
    auditGeneral_(user, eligible ? 'Aplico sello automatico' : 'Retiro sello automatico', 'Pasaporte', passport.id,
      {citaId:citaId, sellos:oldCount}, {citaId:citaId, sellos:newCount}, eligible ? 'Cita atendida pagada' : 'Cita ya no es valida para Pasaporte');
  } catch(auditErr) {}
  return {ok:true, changed:true, passportId:passport.id, sellos:newCount};
}

function passportEnsure_(nombre, telefono, user) {
  try {
    migratePasaportesIfNeeded_();
    if (!nombre) return {ok:false,error:'Falta nombre'};
    var sh = getPasaportesSheet();
    var map = passportHeaderMap_(sh);
    var existing = findPassportByName_(nombre);
    if (existing) {
      if (telefono && map.Telefono) sh.getRange(existing.row, map.Telefono).setValue(telefono);
      existing.telefono = telefono || existing.telefono || '';
      existing.link = passportPublicUrl_(existing.id, existing.token);
      return {ok:true, passport: existing};
    }
    var id = makePassportId_();
    var token = makePassportToken_();
    var row = new Array(sh.getLastColumn()).fill('');
    row[map.nombre - 1] = nombre;
    row[map.passport - 1] = JSON.stringify({manualAdjustment:0});
    row[map.descarga - 1] = '{}';
    row[map.actualizado - 1] = new Date();
    row[map.PasaporteID - 1] = id;
    row[map.TokenPublico - 1] = token;
    row[map.Estado - 1] = 'ACTIVO';
    row[map.FechaCreacionToken - 1] = new Date();
    row[map.FechaUltimoAcceso - 1] = '';
    row[map.Telefono - 1] = telefono || '';
    row[map.CitasAplicadasJSON - 1] = '{}';
    sh.appendRow(row);
    try { auditGeneral_(user, 'Creo pasaporte seguro', 'Pasaporte', id, '', nombre, ''); } catch(auditErr) {}
    return {ok:true, passport: {nombre:nombre, telefono:telefono || '', passport:passportProgressForClient_({manualAdjustment:0}, {}), descarga:{}, id:id, token:token, estado:'ACTIVO', link:passportPublicUrl_(id, token)}};
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

function passportAdminList_() {
  migratePasaportesIfNeeded_();
  var sh = getPasaportesSheet();
  var map = passportHeaderMap_(sh);
  var rows = sh.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < rows.length; i++) list.push(passportRowToObject_(rows[i], map, i + 1));
  return {ok:true, pasaportes:list};
}

function passportSaveProgress_(id, passportJson, descargaJson, user) {
  var sh = getPasaportesSheet();
  var map = passportHeaderMap_(sh);
  var found = findPassportById_(id);
  if (!found) return {ok:false,error:'Pasaporte no encontrado'};
  var row = sh.getRange(found.row, 1, 1, sh.getLastColumn()).getValues()[0];
  var currentPassport = parseJsonSafe_(row[map.passport - 1]);
  var applied = passportAppliedAppointments_(row, map);
  var before = passportEffectiveStampCount_(currentPassport, applied);
  var requested = passportLegacyManualCount_(parseJsonSafe_(passportJson));
  var manualAdjustment = requested - passportActiveAppointmentCount_(applied);
  sh.getRange(found.row, map.passport).setValue(JSON.stringify({manualAdjustment:manualAdjustment}));
  sh.getRange(found.row, map.descarga).setValue(descargaJson || '{}');
  sh.getRange(found.row, map.actualizado).setValue(new Date());
  try { auditGeneral_(user, 'Modifico sellos manualmente', 'Pasaporte', id, {sellos:before}, {sellos:requested, ajusteManual:manualAdjustment}, 'Correccion manual desde panel'); } catch(auditErr) {}
  return {ok:true, passport: findPassportById_(id)};
}

function passportRegenerateToken_(id, user) {
  var sh = getPasaportesSheet();
  var map = passportHeaderMap_(sh);
  var found = findPassportById_(id);
  if (!found) return {ok:false,error:'Pasaporte no encontrado'};
  var token = makePassportToken_();
  sh.getRange(found.row, map.TokenPublico).setValue(token);
  sh.getRange(found.row, map.FechaCreacionToken).setValue(new Date());
  try { auditGeneral_(user, 'Regenero token pasaporte', 'Pasaporte', id, 'Token anterior', 'Token nuevo', 'El enlace anterior queda invalido'); } catch(auditErr) {}
  found.token = token;
  found.link = passportPublicUrl_(id, token);
  return {ok:true, passport: found};
}

function passportDeactivate_(id, user) {
  var sh = getPasaportesSheet();
  var map = passportHeaderMap_(sh);
  var found = findPassportById_(id);
  if (!found) return {ok:false,error:'Pasaporte no encontrado'};
  sh.getRange(found.row, map.Estado).setValue('INACTIVO');
  try { auditGeneral_(user, 'Desactivo pasaporte', 'Pasaporte', id, 'ACTIVO', 'INACTIVO', ''); } catch(auditErr) {}
  return {ok:true, passport:findPassportById_(id)};
}

function passportReactivate_(id, user) {
  var sh = getPasaportesSheet();
  var map = passportHeaderMap_(sh);
  var found = findPassportById_(id);
  if (!found) return {ok:false,error:'Pasaporte no encontrado'};
  sh.getRange(found.row, map.Estado).setValue('ACTIVO');
  try { auditGeneral_(user, 'Reactivo pasaporte', 'Pasaporte', id, 'INACTIVO', 'ACTIVO', ''); } catch(auditErr) {}
  return {ok:true, passport:findPassportById_(id)};
}

function passportBackupAndMigrate_() {
  var ss = getOrCreateSheet();
  var sh = getPasaportesSheet();
  var stamp = Utilities.formatDate(new Date(), 'America/Bogota', 'yyyyMMdd_HHmmss');
  sh.copyTo(ss).setName('BACKUP_Pasaportes_' + stamp);
  var mig = migratePasaportesIfNeeded_();
  return {ok:true, backup:'BACKUP_Pasaportes_' + stamp, migrated:mig.migrated};
}

// -------------------------------------------------------------
//  RECORDATORIOS MENSUALES DE REAGENDAMIENTO
// -------------------------------------------------------------

// Devuelve pacientes cuya Ãºltima cita fue hace ~4 semanas (semana4) o 5+ semanas (semana5)
function getRemindersData() {
  var ss   = getOrCreateSheet();
  var rows = ss.getSheetByName('Citas').getDataRange().getValues();

  // Construir mapa: Ãºltimo registro por paciente
  var map = {};
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (r[10] === 'Cancelada') continue;
    var nombre = ('' + (r[2] || '')).trim();
    var fecha  = sd(r[7]);
    if (!nombre || !fecha) continue;
    var phone = ('' + (r[3] || '')).replace(/\D/g, '');
    var email = ('' + (r[4] || '')).trim();
    var serv  = r[5] || '';
    if (!map[nombre] || fecha > map[nombre].lastFecha) {
      map[nombre] = { nombre: nombre, telefono: phone, email: email, lastFecha: fecha, lastServicio: serv };
    }
  }

  var now = new Date(); now.setHours(0,0,0,0);
  var semana4 = [], semana5 = [];

  for (var key in map) {
    var p = map[key];
    var dp = p.lastFecha.split('-');
    var lastDate = new Date(+dp[0], +dp[1]-1, +dp[2]);
    var dias = Math.floor((now - lastDate) / 86400000);
    if (dias >= 28 && dias < 35) semana4.push({ nombre: p.nombre, telefono: p.telefono, email: p.email, lastServicio: p.lastServicio, lastFecha: p.lastFecha, dias: dias });
    else if (dias >= 35 && dias < 70) semana5.push({ nombre: p.nombre, telefono: p.telefono, email: p.email, lastServicio: p.lastServicio, lastFecha: p.lastFecha, dias: dias });
  }

  semana4.sort(function(a,b){ return a.dias - b.dias; });
  semana5.sort(function(a,b){ return a.dias - b.dias; });
  return { ok: true, semana4: semana4, semana5: semana5 };
}

// Devuelve pacientes cuya Ãºltima cita fue hace 90+ dÃ­as (3 meses o mÃ¡s)
function getInactivosData() {
  try {
    var ss   = getOrCreateSheet();
    var rows = ss.getSheetByName('Citas').getDataRange().getValues();

    // Columnas Citas: [0]ID [1]FechaReg [2]Nombre [3]Telefono [4]Email
    //                 [5]Servicio [6]Modalidad [7]FechaCita [8]Hora [9]Precio [10]Estado
    var map = {};
    for (var i = 1; i < rows.length; i++) {
      var r      = rows[i];
      var estado = ('' + (r[10] || '')).trim();
      if (estado === 'Cancelada') continue;
      if (estado === 'Registro') continue;
      var serv   = ('' + (r[5] || '')).trim();
      if (esRegistro(serv)) continue;
      var nombre = ('' + (r[2] || '')).trim();
      var fecha  = sd(r[7]);
      if (!nombre || !fecha) continue;
      var phone  = ('' + (r[3] || '')).replace(/\D/g, '');
      var email  = ('' + (r[4] || '')).trim();
      var key    = nombre.toLowerCase();
      if (!map[key] || fecha > map[key].lastFecha) {
        map[key] = { nombre: nombre, telefono: phone, email: email, lastServicio: serv, lastFecha: fecha };
      }
    }

    var now = new Date(); now.setHours(0,0,0,0);
    var inactivos = [];
    for (var k in map) {
      var p  = map[k];
      var dp = p.lastFecha.split('-');
      var lastDate = new Date(+dp[0], +dp[1]-1, +dp[2]);
      var dias = Math.floor((now - lastDate) / 86400000);
      if (dias >= 60) {
        inactivos.push({ nombre: p.nombre, telefono: p.telefono, email: p.email, lastServicio: p.lastServicio, lastFecha: p.lastFecha, dias: dias });
      }
    }
    inactivos.sort(function(a,b){ return b.dias - a.dias; });
    return { ok: true, inactivos: inactivos };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// EnvÃ­a emails a todos los pacientes con email registrado que estÃ¡n en semana 4 o 5+
function sendEmailReminders() {
  var data = getRemindersData();
  if (!data.ok) return { ok: false, error: data.error };

  var sent = 0, errors = 0, skipped = 0;
  var all = data.semana4.map(function(p){ return { p:p, semanas:4 }; })
           .concat(data.semana5.map(function(p){ return { p:p, semanas:5 }; }));

  for (var i = 0; i < all.length; i++) {
    var item = all[i];
    var p    = item.p;
    if (!p.email || p.email.indexOf('@') < 0) { skipped++; continue; }

    var primero = p.nombre.split(' ')[0];
    var asunto  = item.semanas === 4
      ? ('â° ' + primero + ', ya es momento de tu prÃ³xima descarga muscular')
      : ('ðŸ’† ' + primero + ', lleva 5 semanas desde tu Ãºltima sesiÃ³n');

    try {
      GmailApp.sendEmail(p.email, asunto, '', {
        htmlBody: buildReminderMensualEmail(p.nombre, item.semanas),
        name: 'Cuidándote Fisioterapia'
      });
      sent++;
    } catch(e) {
      errors++;
      Logger.log('Error email ' + p.email + ': ' + e.message);
    }
  }

  // Resumen para Jessica
  if (sent > 0 || skipped > 0) {
    GmailApp.sendEmail(JESSICA_EMAIL,
      'Recordatorios de reagendamiento enviados â€” ' + sent + ' email(s)',
      'Resumen del envÃ­o automÃ¡tico de recordatorios:\n\n' +
      'âœ… Emails enviados: ' + sent + '\n' +
      'â­ Sin email (WhatsApp manual): ' + skipped + '\n' +
      'âŒ Errores: ' + errors + '\n\n' +
      'Entra al panel admin â†’ Recordatorios para enviarles WhatsApp a los pacientes sin email.');
  }

  return { ok: true, sent: sent, skipped: skipped, errors: errors };
}

function buildReminderMensualEmail(nombre, semanas) {
  var primero = nombre.split(' ')[0];
  var msg = semanas === 4
    ? ('Ya vamos en la <strong>semana 4</strong> desde tu Ãºltima descarga muscular â€” la prÃ³xima semana serÃ­a el momento ideal para hacerla antes de que el cuerpo empiece a acumular tensiÃ³n de nuevo.')
    : ('Ya se cumplieron las <strong>5 semanas</strong> desde tu Ãºltima sesiÃ³n de descarga â€” es el momento de reagendar. Mantener la frecuencia es lo que hace que los resultados se sostengan.');

  return '<div style="font-family:Arial,sans-serif;max-width:520px;margin:0 auto;border:1px solid #e5e7eb;border-radius:12px;overflow:hidden">' +
    '<div style="background:#0d9488;padding:24px 32px;text-align:center">' +
    '<h1 style="color:#fff;margin:0;font-size:19px">â° Es momento de tu prÃ³xima sesiÃ³n</h1>' +
    '<p style="color:#ccfbf1;margin:6px 0 0;font-size:13px">Cuidándote Fisioterapia</p>' +
    '</div>' +
    '<div style="padding:28px 32px">' +
    '<p style="margin:0 0 14px;font-size:15px">Hola <strong>' + primero + '</strong>! ðŸ‘‹ Te escribimos de Cuidándote Fisioterapia.</p>' +
    '<p style="margin:0 0 24px;font-size:15px;line-height:1.65;color:#374151">' + msg + '</p>' +
    '<div style="text-align:center;margin:24px 0">' +
    '<a href="https://jessicaocampoft-ctrl.github.io/#agenda" style="background:#0d9488;color:#fff;padding:14px 32px;border-radius:8px;text-decoration:none;font-weight:600;font-size:15px;display:inline-block">Agendar mi cita ðŸ“…</a>' +
    '</div>' +
    '<hr style="border:none;border-top:1px solid #e5e7eb;margin:24px 0">' +
    '<p style="font-size:13px;color:#6b7280;margin:0 0 12px">Â¿CÃ³mo ha sido tu experiencia? Tu opiniÃ³n me ayuda a mejorar:</p>' +
    '<div style="text-align:center;margin:0 0 20px">' +
    '<a href="https://forms.gle/srX1enyKN59n8TfQA" style="background:#f9fafb;border:1px solid #e5e7eb;color:#0d9488;padding:11px 24px;border-radius:8px;text-decoration:none;font-weight:600;font-size:14px;display:inline-block">â­ Responder encuesta de satisfacciÃ³n</a>' +
    '</div>' +
    '<p style="font-size:13px;color:#6b7280;margin:0">Â¿Prefieres escribirme directamente?<br>' +
    '<a href="https://wa.me/573136467945" style="color:#0d9488">+57 313 646 7945 (WhatsApp)</a></p>' +
    '</div>' +
    '<div style="background:#f9fafb;padding:16px 32px;text-align:center;font-size:12px;color:#9ca3af">' +
    'Cuidándote Fisioterapia Â· Pereira, Colombia<br>' +
    '<a href="https://jessicaocampoft-ctrl.github.io" style="color:#0d9488">jessicaocampoft-ctrl.github.io</a>' +
    '</div></div>';
}

// Ejecuta ESTA funciÃ³n UNA sola vez para activar el trigger semanal automÃ¡tico:
function setupReminderTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'autoSendReminders') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('autoSendReminders')
    .timeBased().everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8).inTimezone('America/Bogota').create();
  Logger.log('Trigger activado: recordatorios todos los lunes a las 8am (hora Colombia).');
}

function autoSendReminders() {
  var result = sendEmailReminders();
  Logger.log('autoSendReminders: enviados=' + result.sent + ', sinEmail=' + result.skipped + ', errores=' + result.errors);
}

// â”€â”€ RESEÃ‘AS GOOGLE â”€â”€
function getGoogleReviews() {
  var PLACE_ID = 'ChIJVwU1iJ15sCARAQ_jFCdVsXI';
  var API_KEY  = 'AIzaSyAKtsK8EaAG0GE_0Ma-mNoaMwy1ZG0gEv8';
  var CACHE_KEY = 'google_reviews_cache_v1';
  var CACHE_TS_KEY = 'google_reviews_cache_v1_ts';
  var CACHE_MAX_AGE_MS = 24 * 60 * 60 * 1000;
  try {
    var props = PropertiesService.getScriptProperties();
    var cached = props.getProperty(CACHE_KEY);
    var cachedTs = Number(props.getProperty(CACHE_TS_KEY) || 0);
    if (cached && cachedTs && (Date.now() - cachedTs) < CACHE_MAX_AGE_MS) {
      return { ok: true, data: JSON.parse(cached), cached: true };
    }

    var url = 'https://maps.googleapis.com/maps/api/place/details/json'
      + '?place_id=' + encodeURIComponent(PLACE_ID)
      + '&fields=rating,user_ratings_total,reviews'
      + '&language=es'
      + '&key=' + encodeURIComponent(API_KEY);
    var res  = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true
    });
    var data = JSON.parse(res.getContentText());
    if (data.status !== 'OK') {
      if (cached) {
        return { ok: true, data: JSON.parse(cached), cached: true, stale: true };
      }
      return {
        ok: false,
        error: data.error_message || data.status || 'No fue posible cargar reseÃ±as reales de Google'
      };
    }
    var result = data.result || {};
    var payload = {
      rating: result.rating || 0,
      userRatingCount: result.user_ratings_total || 0,
      reviews: (result.reviews || []).map(function(r) {
        return {
          rating: r.rating || 0,
          text: { text: r.text || '' },
          originalText: { text: r.text || '' },
          authorAttribution: {
            displayName: r.author_name || 'Paciente',
            photoUri: r.profile_photo_url || ''
          },
          relativePublishTimeDescription: r.relative_time_description || ''
        };
      })
    };
    props.setProperty(CACHE_KEY, JSON.stringify(payload));
    props.setProperty(CACHE_TS_KEY, String(Date.now()));
    return { ok: true, data: payload, cached: false };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// =============================================================
//  EVALUACIÃ“N EXPRESS CROSSFIT â€” GeneraciÃ³n de reporte con IA
// =============================================================

function generateEvalReport(d, photos) {
  photos = photos || {};
  try {
    if (!GEMINI_API_KEY || GEMINI_API_KEY === 'PEGA_AQUI_TU_CLAVE_GEMINI') {
      return { ok: false, error: 'Configura GEMINI_API_KEY en el script. Obtenla gratis en aistudio.google.com' };
    }

    var parts = [{ text: buildEvalPrompt(d) }];

    // AÃ±adir fotos posturales y de tests a la solicitud multimodal
    var photoLabels = {
      frontal:  'FOTO POSTURAL FRONTAL â€” analiza alineaciÃ³n de cabeza, hombros, crestas ilÃ­acas, rodillas y pies',
      lateral:  'FOTO POSTURAL LATERAL â€” analiza adelantamiento cefÃ¡lico, cifosis, lordosis, posiciÃ³n rodilla',
      posterior:'FOTO POSTURAL POSTERIOR â€” analiza escoliosis, asimetrÃ­a escapular, pies y talones',
      ds:       'FOTO DEEP SQUAT â€” analiza profundidad, posiciÃ³n de rodillas, talones y tronco',
      oh:       'FOTO OVERHEAD REACH â€” analiza contacto de manos con pared y compensaciones',
      slsd:     'FOTO SINGLE LEG SQUAT DERECHO â€” analiza control de rodilla y cadera',
      slsi:     'FOTO SINGLE LEG SQUAT IZQUIERDO â€” analiza control de rodilla y cadera',
      shd:      'FOTO SHOULDER CLEARING MANO D ARRIBA â€” analiza distancia entre puÃ±os',
      shi:      'FOTO SHOULDER CLEARING MANO I ARRIBA â€” analiza distancia entre puÃ±os',
      bmd:      'FOTO BALANCE MONOPODAL DERECHO â€” analiza postura y estrategia de equilibrio',
      bmi:      'FOTO BALANCE MONOPODAL IZQUIERDO â€” analiza postura y estrategia de equilibrio',
      trdd:     'FOTO TRENDELENBURG APOYO DERECHO â€” analiza nivel pÃ©lvico y caÃ­da contralateral',
      trdi:     'FOTO TRENDELENBURG APOYO IZQUIERDO â€” analiza nivel pÃ©lvico y caÃ­da contralateral',
      sbd:      'FOTO SHIN BOX PIERNA DERECHA â€” analiza posiciÃ³n de shin y rango de movimiento',
      sbi:      'FOTO SHIN BOX PIERNA IZQUIERDA â€” analiza posiciÃ³n de shin y rango de movimiento',
      bdd:      'FOTO BIRD DOG LADO DERECHO â€” analiza neutralidad lumbar y control rotacional',
      bdi:      'FOTO BIRD DOG LADO IZQUIERDO â€” analiza neutralidad lumbar y control rotacional',
      dbd:      'FOTO DEAD BUG LADO DERECHO â€” analiza neutro lumbar al extender extremidades',
      dbi:      'FOTO DEAD BUG LADO IZQUIERDO â€” analiza neutro lumbar al extender extremidades',
      pk:       'FOTO PLANK â€” analiza alineaciÃ³n de cadera, espalda y posiciÃ³n general'
    };

    var photoCount = 0;
    for (var key in photoLabels) {
      if (photos[key]) {
        var base64 = photos[key].replace(/^data:image\/[^;]+;base64,/, '');
        parts.push({ text: '\n[' + photoLabels[key] + ']' });
        parts.push({ inlineData: { mimeType: 'image/jpeg', data: base64 } });
        photoCount++;
      }
    }

    if (photoCount > 0) {
      parts.push({ text: '\nCon base en las ' + photoCount + ' fotografÃ­as anteriores y los datos ingresados, incluye en tu anÃ¡lisis observaciones especÃ­ficas de lo que ves en las imÃ¡genes de cada test. Valida o corrige los scores asignados por la fisioterapeuta si lo consideras necesario, explicando el motivo.' });
    }

    var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=' + GEMINI_API_KEY;
    var response = UrlFetchApp.fetch(url, {
      method:      'post',
      contentType: 'application/json',
      muteHttpExceptions: true,
      payload: JSON.stringify({
        contents: [{ parts: parts }],
        generationConfig: { temperature: 0.7, maxOutputTokens: 1500 }
      })
    });

    var result = JSON.parse(response.getContentText());
    if (result.error) return { ok: false, error: result.error.message };

    var text = result.candidates[0].content.parts[0].text;
    return { ok: true, report: text };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// =============================================================
//  EVENTOS (agenda interna)
// =============================================================

function getEventosSheet() {
  var ss = getOrCreateSheet();
  var sh = ss.getSheetByName('Eventos');
  if (!sh) {
    sh = ss.insertSheet('Eventos');
    sh.appendRow(['ID','TÃ­tulo','Tipo','Fecha','HoraInicio','HoraFin','DuraciÃ³n','Cobro','Notas']);
    sh.getRange(1,1,1,9).setFontWeight('bold').setBackground('#7c3aed').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}

function crearEvento(p) {
  var d  = JSON.parse(p.data);
  var sh = getEventosSheet();
  var id = 'EVT-' + new Date().getTime();
  sh.appendRow([id, d.titulo, d.tipo, d.fecha, d.horaInicio, d.horaFin, d.duracion || '', d.cobro || 'Sin cobro', d.notas || '']);
  // Forzar fecha y horas como texto para evitar auto-detecciÃ³n de Sheets
  var lastRow = sh.getLastRow();
  sh.getRange(lastRow, 4, 1, 3).setNumberFormat('@');
  return {ok: true, id: id};
}

function eliminarEvento(p) {
  var sh   = getEventosSheet();
  var rows = sh.getDataRange().getValues();
  for (var i = rows.length - 1; i >= 1; i--) {
    if ('' + rows[i][0] === '' + p.id) {
      sh.deleteRow(i + 1);
      return {ok: true};
    }
  }
  return {ok: false, error: 'Evento no encontrado'};
}

// =============================================================
//  SISTEMA DE CÃ“DIGOS â€” REF-MES-NNN  /  BONO-MES-NNN
// =============================================================

function getCodigosSheet() {
  var ss = getOrCreateSheet();
  var sh = ss.getSheetByName('Codigos');
  if (!sh) {
    sh = ss.insertSheet('Codigos');
    sh.appendRow(['CÃ³digo','Tipo','Paciente','TelÃ©fono','Referido por','Fecha','Estado','CÃ³digoRef']);
    sh.getRange(1,1,1,8).setFontWeight('bold').setBackground('#1BBFB0').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}

function generarCodigo(p) {
  var tipo = p.tipo || 'REF';
  var sh   = getCodigosSheet();
  var MESES = ['ENE','FEB','MAR','ABR','MAY','JUN','JUL','AGO','SEP','OCT','NOV','DIC'];
  var mes    = MESES[new Date().getMonth()];
  var prefix = tipo + '-' + mes + '-';

  var datos  = sh.getDataRange().getValues();
  var maxNum = 0;
  for (var i = 1; i < datos.length; i++) {
    var cod = '' + (datos[i][0] || '');
    if (cod.indexOf(prefix) === 0) {
      var num = parseInt(cod.replace(prefix, ''), 10) || 0;
      if (num > maxNum) maxNum = num;
    }
  }

  var codigo = prefix + ('' + (maxNum + 1)).padStart(3, '0');
  return {ok: true, codigo: codigo};
}

function registrarCodigo(p) {
  var data = JSON.parse(p.data);
  // data: { codigo, tipo, paciente, telefono, referidoPor, codigoRef }
  var sh    = getCodigosSheet();
  var fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  sh.appendRow([
    data.codigo,
    data.tipo,
    data.paciente    || '',
    ('' + (data.telefono || '')).replace(/\D/g, ''),
    data.referidoPor || '',
    fecha,
    'Activo',
    data.codigoRef   || ''
  ]);
  // Formatear telÃ©fono como texto
  sh.getRange(sh.getLastRow(), 4).setNumberFormat('@');
  return {ok: true, codigo: data.codigo};
}

function actualizarCodigo(p) {
  var sh    = getCodigosSheet();
  var rows  = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (('' + rows[i][0]) === p.codigo) {
      sh.getRange(i + 1, 7).setValue(p.estado); // columna 7 = Estado
      return {ok: true};
    }
  }
  return {ok: false, error: 'CÃ³digo no encontrado'};
}

function getCodigos() {
  var sh    = getCodigosSheet();
  var rows  = sh.getDataRange().getValues();
  var lista = [];
  for (var i = 1; i < rows.length; i++) {
    lista.push({
      codigo:      '' + (rows[i][0] || ''),
      tipo:        '' + (rows[i][1] || ''),
      paciente:    '' + (rows[i][2] || ''),
      telefono:    '' + (rows[i][3] || ''),
      referidoPor: '' + (rows[i][4] || ''),
      fecha:       '' + (rows[i][5] || ''),
      estado:      '' + (rows[i][6] || ''),
      codigoRef:   '' + (rows[i][7] || '')
    });
  }
  return {ok: true, codigos: lista};
}

// Elimina citas sin hora de la hoja Citas
function cleanCitasSinHora() {
  var sheet = getOrCreateSheet().getSheetByName('Citas');
  var rows  = sheet.getDataRange().getValues();
  var deleted = 0;
  for (var i = rows.length - 1; i >= 1; i--) {
    var hora = rows[i][8]; // columna Hora (Ã­ndice 8)
    var horaStr = st(hora);
    if (!horaStr || horaStr === '0:0' || isMidnightBookingTime_(horaStr)) {
      sheet.deleteRow(i + 1);
      deleted++;
    }
  }
  return { ok: true, deleted: deleted };
}

// Elimina citas recientes/futuras que quedaron por fuera de la jornada permitida
// (por ejemplo 21:00). No toca registros ni citas canceladas.
function cleanInvalidCitaTimes() {
  var sheet = getOrCreateSheet().getSheetByName('Citas');
  var rows  = sheet.getDataRange().getValues();
  var deleted = 0;
  var items = [];
  var now = new Date();
  var yesterday = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1);
  var cutoff = fmtDate(yesterday);

  for (var i = rows.length - 1; i >= 1; i--) {
    var r = rows[i];
    var id = '' + (r[0] || '');
    var nombre = '' + (r[2] || '');
    var servicio = '' + (r[5] || '');
    var modalidad = '' + (r[6] || '');
    var fecha = sd(r[7]);
    var hora = st(r[8]);
    var estado = '' + (r[10] || '');
    if (!bookingIsActive_(estado, servicio)) continue;
    if (!fecha || fecha < cutoff) continue;
    var check = validateBookingSchedule_(fecha, hora, servicio, modalidad);
    if (!check.ok) {
      try {
        var dp = fecha.split('-');
        var dayStart = new Date(+dp[0], +dp[1]-1, +dp[2], 0, 0, 0);
        var dayEnd = new Date(+dp[0], +dp[1]-1, +dp[2], 23, 59, 59);
        var calEvs = CalendarApp.getDefaultCalendar().getEvents(dayStart, dayEnd);
        for (var k = 0; k < calEvs.length; k++) {
          var title = calEvs[k].getTitle() || '';
          var evTime = pad(calEvs[k].getStartTime().getHours()) + ':' + pad(calEvs[k].getStartTime().getMinutes());
          if (title.indexOf('[CITA]') === 0 && title.indexOf(nombre) > -1 && evTime === hora) {
            calEvs[k].deleteEvent();
            break;
          }
        }
      } catch(x) {}
      sheet.deleteRow(i + 1);
      deleted++;
      items.push({id: id, nombre: nombre, fecha: fecha, hora: hora, motivo: check.error});
    }
  }
  return { ok: true, deleted: deleted, items: items };
}

function buildEvalPrompt(d) {
  var hallazgos = (d.hallazgos && d.hallazgos.length)
    ? d.hallazgos.join(', ')
    : 'Sin hallazgos registrados';

  var screensText = '';
  if (d.screens) {
    var sc = d.screens;
    screensText += '\nRESULTADOS PRUEBAS FUNCIONALES (escala 0-3, donde 3=normal, 0=dolor/no puede):\n';
    if (sc.deepSquat) {
      screensText += 'â€¢ Deep Squat: ' + (sc.deepSquat.score !== null ? sc.deepSquat.score + '/3' : 'No evaluado');
      if (sc.deepSquat.obs) screensText += ' â€” ' + sc.deepSquat.obs;
      screensText += '\n';
    }
    if (sc.overheadReach) {
      screensText += 'â€¢ Overhead Reach: ' + (sc.overheadReach.score !== null ? sc.overheadReach.score + '/3' : 'No evaluado');
      if (sc.overheadReach.obs) screensText += ' â€” ' + sc.overheadReach.obs;
      screensText += '\n';
    }
    if (sc.singleLegSquat) {
      screensText += 'â€¢ Single Leg Squat: D=' + (sc.singleLegSquat.scoreD !== null ? sc.singleLegSquat.scoreD + '/3' : 'No eval') +
        ', I=' + (sc.singleLegSquat.scoreI !== null ? sc.singleLegSquat.scoreI + '/3' : 'No eval');
      if (sc.singleLegSquat.obs) screensText += ' â€” ' + sc.singleLegSquat.obs;
      screensText += '\n';
    }
    if (sc.shoulderMob) {
      screensText += 'â€¢ Shoulder Clearing: D=' + (sc.shoulderMob.scoreD !== null ? sc.shoulderMob.scoreD + '/3' : 'No eval') +
        ', I=' + (sc.shoulderMob.scoreI !== null ? sc.shoulderMob.scoreI + '/3' : 'No eval');
      if (sc.shoulderMob.obs) screensText += ' â€” ' + sc.shoulderMob.obs;
      screensText += '\n';
    }
    if (sc.trendelenburg) {
      screensText += 'â€¢ Trendelenburg: D=' + (sc.trendelenburg.resultD || 'No eval') +
        ', I=' + (sc.trendelenburg.resultI || 'No eval');
      if (sc.trendelenburg.obs) screensText += ' â€” ' + sc.trendelenburg.obs;
      screensText += '\n';
    }
    if (sc.shinBox) {
      screensText += 'â€¢ Shin Box: D RI=' + (sc.shinBox.riD || '?') + 'Â° RE=' + (sc.shinBox.reD || '?') +
        'Â° Trans=' + (sc.shinBox.transD !== null && sc.shinBox.transD !== undefined ? sc.shinBox.transD + '/3' : 'No eval') +
        ' | I RI=' + (sc.shinBox.riI || '?') + 'Â° RE=' + (sc.shinBox.reI || '?') +
        'Â° Trans=' + (sc.shinBox.transI !== null && sc.shinBox.transI !== undefined ? sc.shinBox.transI + '/3' : 'No eval');
      if (sc.shinBox.obs) screensText += ' â€” ' + sc.shinBox.obs;
      screensText += ' (Normal RI 35-45Â°, RE 40-60Â°)\n';
    }
    if (sc.birdDog) {
      screensText += 'â€¢ Bird Dog (control core): D=' + (sc.birdDog.scoreD !== null && sc.birdDog.scoreD !== undefined ? sc.birdDog.scoreD + '/3' : 'No eval') +
        ', I=' + (sc.birdDog.scoreI !== null && sc.birdDog.scoreI !== undefined ? sc.birdDog.scoreI + '/3' : 'No eval');
      if (sc.birdDog.obs) screensText += ' â€” ' + sc.birdDog.obs;
      screensText += '\n';
    }
  }

  return 'Eres una fisioterapeuta deportiva especializada en CrossFit llamada Jessica Ocampo, con sede en Pereira, Colombia.\n' +
    'Genera un reporte de evaluaciÃ³n postural express profesional, empÃ¡tico y motivador para:\n\n' +
    'ATLETA: ' + d.nombre + (d.edad && d.edad !== 'N/A' ? ', ' + d.edad + ' aÃ±os' : '') + '\n' +
    'NIVEL CROSSFIT: ' + (d.nivel || 'No especificado') + '\n' +
    'TIEMPO EN CROSSFIT: ' + (d.anios || 'No especificado') + '\n' +
    'OBJETIVO: ' + (d.objetivo || 'No especificado') + '\n' +
    'MOLESTIAS ACTUALES: ' + (d.molestias || 'Ninguna') + '\n' +
    'SEVERIDAD ESTIMADA: ' + (d.severidad || 'No especificada') + '\n' +
    'HALLAZGOS POSTURALES: ' + hallazgos + '\n' +
    screensText +
    (d.observaciones ? 'OBSERVACIONES ADICIONALES: ' + d.observaciones + '\n' : '') +
    '\nGenera DOS reportes separados. El primero es tÃ©cnico (para la fisioterapeuta), el segundo es para el paciente (lenguaje simple). SepÃ¡ralos exactamente con la lÃ­nea: ===PACIENTE===\n\n' +

    'â”â”â” REPORTE TÃ‰CNICO (para la fisioterapeuta) â”â”â”\n' +
    'Usa terminologÃ­a clÃ­nica. Secciones con tÃ­tulos en negrilla:\n\n' +
    '**RESUMEN EJECUTIVO**\n' +
    '[2-3 oraciones: estado postural y funcional general, impacto en rendimiento CrossFit. Si hay fotos, describe hallazgos visuales relevantes.]\n\n' +
    '**ANÃLISIS VISUAL DE FOTOS**\n' +
    '[SOLO si se enviaron fotos: hallazgos por foto. Si no hay fotos, omite esta secciÃ³n.]\n\n' +
    '**ANÃLISIS POR ZONAS**\n' +
    '[Por cada zona con hallazgos: significado biomecÃ¡nico y afectaciÃ³n en movimientos CrossFit (snatch, clean, squat, deadlift). Omite zonas sin hallazgos.]\n\n' +
    '**ANÃLISIS FUNCIONAL**\n' +
    '[Interpreta Trendelenburg (glÃºteo medio), Shin Box (movilidad cadera), Bird Dog (control core). Correlaciona con hallazgos posturales.]\n\n' +
    '**RIESGOS IDENTIFICADOS**\n' +
    '[MÃ¡ximo 4 riesgos concretos de lesiÃ³n. EspecÃ­ficos para CrossFit. ViÃ±etas con â€¢]\n\n' +
    '**PLAN DE ACCIÃ“N RECOMENDADO**\n' +
    '[3-5 recomendaciones clÃ­nicas por prioridad: tipo de intervenciÃ³n en fisioterapia (movilidad, estabilizaciÃ³n, correcciÃ³n postural). ViÃ±etas con â€¢]\n\n' +

    '===PACIENTE===\n\n' +

    'â”â”â” REPORTE PARA EL PACIENTE â”â”â”\n' +
    'Lenguaje simple, cÃ¡lido y motivador. Sin tÃ©rminos clÃ­nicos. Secciones:\n\n' +
    '**Â¿QUÃ‰ ENCONTRAMOS HOY?**\n' +
    '[2-3 oraciones explicando en palabras simples lo que se encontrÃ³ y cÃ³mo afecta su entrenamiento. Sin jerga mÃ©dica.]\n\n' +
    '**LO QUE ESTÃ BIEN ðŸ’ª**\n' +
    '[1-2 oraciones destacando los puntos fuertes del atleta. Siempre hay algo positivo.]\n\n' +
    '**RECOMENDACIONES PARA TU ENTRENAMIENTO**\n' +
    '[1 sola recomendaciÃ³n general, en 2-3 oraciones. Sin mencionar ejercicios especÃ­ficos ni movimientos de CrossFit por nombre. Enfocada en un principio general (ej: tÃ©cnica sobre velocidad, escuchar el cuerpo, etc.). Lenguaje simple.]\n\n' +
    '**TU SIGUIENTE PASO**\n' +
    '[1 pÃ¡rrafo corto y motivador invitÃ¡ndolo a agendar su plan de fisioterapia. Menciona que ya identificaste exactamente quÃ© trabajar y que los resultados se ven rÃ¡pido con un plan personalizado. CÃ¡lido y sin presiÃ³n.]\n\n' +

    'REGLAS GLOBALES: Reporte tÃ©cnico mÃ¡x 500 palabras. Reporte paciente mÃ¡x 300 palabras. Tono tÃ©cnico: preciso y profesional. Tono paciente: cercano, claro, motivador.';
}

// =============================================================
//  ENCUESTA DE SATISFACCIÃ“N â€” NPS y % respuestas desde Google Forms
// =============================================================
// Ejecuta esta funciÃ³n UNA vez para autorizar el permiso de Formularios
function autorizarFormularios() {
  FormApp.openById('1UxoEq1x4GXaG9ghBQJO_C85p3ZPU3T7zeKhy0Ij-UA4');
  Logger.log('AutorizaciÃ³n concedida');
}

function getEncuestaStats_() {
  try {
    var FORM_ID = '1UxoEq1x4GXaG9ghBQJO_C85p3ZPU3T7zeKhy0Ij-UA4';
    var form = FormApp.openById(FORM_ID);
    var responses = form.getResponses();
    var now = new Date();
    var year = now.getFullYear(), month = now.getMonth();

    var mesRes = responses.filter(function(r) {
      var d = r.getTimestamp();
      return d.getFullYear() === year && d.getMonth() === month;
    });

    // Busca la pregunta NPS por palabras clave en el tÃ­tulo (recomiendes/probable)
    // para no confundirla con otras preguntas de calificaciÃ³n del formulario
    var items = form.getItems();
    var npsItem = null;

    for (var i = 0; i < items.length; i++) {
      var titulo = items[i].getTitle().toLowerCase();
      var esNPS = titulo.indexOf('recomiendes') > -1 || titulo.indexOf('recomiend') > -1 ||
                  titulo.indexOf('probable') > -1 || titulo.indexOf('nps') > -1;
      if (!esNPS) continue;
      var t = items[i].getType();
      if (t === FormApp.ItemType.LINEAR_SCALE || t === FormApp.ItemType.MULTIPLE_CHOICE) {
        npsItem = items[i]; break;
      }
    }

    // Escala 0-5: 5=Promotor, 4=Pasivo, 0-3=Detractor
    var promotores = 0, pasivos = 0, detractores = 0;
    if (npsItem) {
      var npsId = npsItem.getId();
      mesRes.forEach(function(r) {
        var ir = r.getItemResponses();
        for (var j = 0; j < ir.length; j++) {
          if (ir[j].getItem().getId() === npsId) {
            var score = parseInt(('' + ir[j].getResponse()).trim(), 10);
            // Si parseInt falla (opciones son texto puro sin nÃºmero), mapear por etiqueta
            if (isNaN(score)) {
              var txt = ('' + ir[j].getResponse()).toLowerCase()
                .replace(/[Ã¡Ã Ã¢]/g,'a').replace(/[Ã©Ã¨Ãª]/g,'e')
                .replace(/[Ã­Ã¬Ã®]/g,'i').replace(/[Ã³Ã²Ã´]/g,'o').replace(/[ÃºÃ¹Ã»]/g,'u');
              if (txt.indexOf('totalmente') > -1)      score = 5;
              else if (txt.indexOf('muy') > -1)         score = 4;
              else if (txt.indexOf('medianamente') > -1) score = 2;
              else if (txt.indexOf('poco') > -1)        score = 1;
              else if (txt.indexOf('nada') > -1)        score = 0;
              else if (txt.trim() === 'probable')       score = 3;
            }
            if (!isNaN(score)) {
              if (score === 5)      promotores++;
              else if (score === 4) pasivos++;
              else                  detractores++;
            }
            break;
          }
        }
      });
    }

    var total = mesRes.length;

    // Recolectar muestra de respuestas reales para diagnÃ³stico
    var rawSample = [];
    if (npsItem && mesRes.length > 0) {
      var npsIdS = npsItem.getId();
      for (var s = 0; s < Math.min(3, mesRes.length); s++) {
        var irS = mesRes[s].getItemResponses();
        for (var q = 0; q < irS.length; q++) {
          if (irS[q].getItem().getId() === npsIdS) {
            rawSample.push('' + irS[q].getResponse());
            break;
          }
        }
      }
    }

    return {
      ok: true,
      totalMes:    total,
      promotores:  promotores,
      pasivos:     pasivos,
      detractores: detractores,
      nps: (npsItem && total > 0) ? Math.round((promotores / total - detractores / total) * 100) : null,
      npsItemEncontrado: npsItem ? npsItem.getTitle() : 'NO ENCONTRADO',
      rawSample: rawSample
    };
  } catch(e) { return { ok: false, error: e.toString() }; }
}

// Ejecuta esta funciÃ³n en GAS para diagnosticar â€” el resultado aparece en "Registro de ejecuciÃ³n"
function debugEncuesta() {
  var result = getEncuestaStats_();
  var FORM_ID = '1UxoEq1x4GXaG9ghBQJO_C85p3ZPU3T7zeKhy0Ij-UA4';
  var form = FormApp.openById(FORM_ID);
  var items = form.getItems();
  items.forEach(function(it) {
    console.log('Item: ' + it.getTitle() + ' | Tipo: ' + it.getType().toString());
  });
  console.log('Stats: ' + JSON.stringify(result));
}

// =============================================================
//  MOTOR CENTRAL DE AUTOMATIZACIONES
// =============================================================

var AUTOMATION_DEFAULTS = {
  emailReminders: true,
  whatsappQueue: true,
  followups: true,
  autoFollowupEmail: false,
  inactivePatients: true,
  paymentAlerts: true,
  dataQuality: true,
  weeklyReport: true,
  backups: true,
  kpiSnapshots: true,
  waitlistMatching: true
};

function getAutomationConfig() {
  var raw = PropertiesService.getScriptProperties().getProperty('AUTOMATION_CONFIG');
  var saved = {};
  try { saved = raw ? JSON.parse(raw) : {}; } catch(e) {}
  var cfg = {};
  for (var k in AUTOMATION_DEFAULTS) cfg[k] = saved[k] === undefined ? AUTOMATION_DEFAULTS[k] : !!saved[k];
  return cfg;
}

function saveAutomationConfig(data) {
  try {
    var incoming = JSON.parse(decodeURIComponent(data || '{}'));
    var cfg = getAutomationConfig();
    for (var k in AUTOMATION_DEFAULTS) if (incoming[k] !== undefined) cfg[k] = !!incoming[k];
    PropertiesService.getScriptProperties().setProperty('AUTOMATION_CONFIG', JSON.stringify(cfg));
    automationLog_('config', 'ConfiguraciÃ³n actualizada', 'ok');
    return {ok:true, config:cfg};
  } catch(e) { return {ok:false, error:e.message}; }
}

function getAutomationSheet_(name, headers) {
  var ss = getOrCreateSheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
  }
  return sh;
}

function automationLog_(job, detail, status) {
  try {
    getAutomationSheet_('AutomationLog', ['timestamp','job','detail','status'])
      .appendRow([new Date(), job, detail, status || 'ok']);
  } catch(e) { Logger.log('automationLog: ' + e.message); }
}

function automationQueueSheet_() {
  return getAutomationSheet_('ColaMensajes', ['id','tipo','nombre','telefono','mensaje','motivo','estado','creado','relacionado']);
}

function queueAutomationMessage_(type, name, phone, message, reason, relatedId, uniqueKey) {
  if (!getAutomationConfig().whatsappQueue) return false;
  var sh = automationQueueSheet_();
  var rows = sh.getDataRange().getValues();
  var key = uniqueKey || [type, relatedId, fmtDate(new Date())].join('|');
  for (var i=1;i<rows.length;i++) if ((''+rows[i][8]) === key) return false;
  var cleanPhone = (''+(phone||'')).replace(/\D/g,'');
  if (cleanPhone && cleanPhone.length <= 10) cleanPhone = '57' + cleanPhone;
  sh.appendRow(['MSG-'+new Date().getTime()+'-'+Math.floor(Math.random()*999), type, name||'', cleanPhone, message||'', reason||'', 'Pendiente', new Date(), key]);
  return true;
}

function getAutomationQueue(status) {
  var rows = automationQueueSheet_().getDataRange().getValues();
  var out = [];
  for (var i=1;i<rows.length;i++) {
    if (!rows[i][0]) continue;
    var item = {id:''+rows[i][0],type:''+rows[i][1],nombre:''+rows[i][2],telefono:''+rows[i][3],mensaje:''+rows[i][4],motivo:''+rows[i][5],estado:''+rows[i][6],creado:rows[i][7] instanceof Date?rows[i][7].toISOString():''+rows[i][7]};
    if (!status || status === 'all' || item.estado.toLowerCase() === status.toLowerCase() || (status === 'pending' && item.estado === 'Pendiente')) out.push(item);
  }
  return {ok:true, items:out.reverse().slice(0,200)};
}

function markAutomationQueueDone(id) {
  var sh = automationQueueSheet_(), rows = sh.getDataRange().getValues();
  for (var i=1;i<rows.length;i++) if ((''+rows[i][0]) === (''+id)) { sh.getRange(i+1,7).setValue('Enviado'); return {ok:true}; }
  return {ok:false,error:'Mensaje no encontrado'};
}

function setupAllAutomations() {
  var handlers = ['runAutomationMorning','runAutomationNight','runAutomationWeekly','expireTemporaryReservations_'];
  ScriptApp.getProjectTriggers().forEach(function(t){ if (handlers.indexOf(t.getHandlerFunction())>=0 || ['sendReminders','autoMarcarAtendidas','autoSendReminders'].indexOf(t.getHandlerFunction())>=0) ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('runAutomationMorning').timeBased().everyDays(1).atHour(7).inTimezone('America/Bogota').create();
  ScriptApp.newTrigger('runAutomationNight').timeBased().everyDays(1).atHour(22).inTimezone('America/Bogota').create();
  ScriptApp.newTrigger('runAutomationWeekly').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).inTimezone('America/Bogota').create();
  ScriptApp.newTrigger('expireTemporaryReservations_').timeBased().everyMinutes(15).create();
  automationLog_('setup','Triggers instalados: diario 7am, diario 10pm, lunes 8am y vencimiento seguro de reservas cada 15 minutos','ok');
  return getAutomationStatus();
}

function getAutomationStatus() {
  var triggers = ScriptApp.getProjectTriggers().map(function(t){ return t.getHandlerFunction(); });
  var logRows = getAutomationSheet_('AutomationLog', ['timestamp','job','detail','status']).getDataRange().getValues();
  var logs=[];
  for(var i=Math.max(1,logRows.length-15);i<logRows.length;i++) logs.push({timestamp:logRows[i][0] instanceof Date?logRows[i][0].toISOString():''+logRows[i][0],job:''+logRows[i][1],detail:''+logRows[i][2],status:''+logRows[i][3]});
  var pending = getAutomationQueue('pending').items.length;
  return {ok:true,config:getAutomationConfig(),triggers:triggers,active:triggers.indexOf('runAutomationMorning')>=0&&triggers.indexOf('runAutomationNight')>=0,pending:pending,logs:logs.reverse()};
}

function runAutomationNow(job) {
  if (job === 'night') return runAutomationNight();
  if (job === 'weekly') return runAutomationWeekly();
  if (job === 'snapshot') return saveKPISnapshot_();
  if (job === 'backup') return createSpreadsheetBackup_();
  return runAutomationMorning();
}

function runAutomationMorning() {
  var cfg=getAutomationConfig(), results={ok:true,job:'morning'};
  var props=PropertiesService.getScriptProperties(), todayKey=fmtDate(new Date());
  if(props.getProperty('AUTO_LAST_MORNING')===todayKey)return{ok:true,job:'morning',skipped:true,reason:'Ya se ejecutÃ³ hoy'};
  try { if(cfg.emailReminders) { sendReminders(); results.reminders=true; } } catch(e) { results.remindersError=e.message; }
  try { if(cfg.followups) results.followups=queuePostSessionFollowups_(); } catch(e2) { results.followupsError=e2.message; }
  try { if(cfg.paymentAlerts) results.paymentAlerts=sendPendingPaymentsSummary_(); } catch(e3) { results.paymentError=e3.message; }
  automationLog_('morning',JSON.stringify(results),results.remindersError?'error':'ok');
  props.setProperty('AUTO_LAST_MORNING',todayKey);
  return results;
}

function runAutomationNight() {
  var cfg=getAutomationConfig(), results={ok:true,job:'night'};
  try { results.attended=autoMarcarAtendidas(); } catch(e) { results.attendedError=e.message; }
  try { if(cfg.kpiSnapshots) results.snapshot=saveKPISnapshot_(); } catch(e2) { results.snapshotError=e2.message; }
  automationLog_('night',JSON.stringify(results),results.attendedError?'error':'ok');
  return results;
}

function runAutomationWeekly() {
  var cfg=getAutomationConfig(), results={ok:true,job:'weekly'};
  var props=PropertiesService.getScriptProperties(), now=new Date(), weekKey=now.getFullYear()+'-W'+Math.ceil((((now-new Date(now.getFullYear(),0,1))/86400000)+new Date(now.getFullYear(),0,1).getDay()+1)/7);
  if(props.getProperty('AUTO_LAST_WEEKLY')===weekKey)return{ok:true,job:'weekly',skipped:true,reason:'Ya se ejecutÃ³ esta semana'};
  try { if(cfg.inactivePatients) { results.inactive=sendEmailReminders(); results.inactiveQueue=queueInactiveReminders_(); } } catch(e) { results.inactiveError=e.message; }
  try { if(cfg.weeklyReport) results.report=sendWeeklyManagementReport_(); } catch(e2) { results.reportError=e2.message; }
  try { if(cfg.backups) results.backup=createSpreadsheetBackup_(); } catch(e3) { results.backupError=e3.message; }
  try { if(cfg.dataQuality) results.quality=sendDataQualitySummary_(); } catch(e4) { results.qualityError=e4.message; }
  automationLog_('weekly',JSON.stringify(results),(results.reportError||results.backupError)?'error':'ok');
  props.setProperty('AUTO_LAST_WEEKLY',weekKey);
  return results;
}

function queuePostSessionFollowups_() {
  var rows=getOrCreateSheet().getSheetByName('Citas').getDataRange().getValues();
  var today=new Date(), queued=0, cfg=getAutomationConfig();
  for(var i=1;i<rows.length;i++) {
    var r=rows[i], status=(''+(r[10]||'')).toLowerCase();
    if(status!=='atendida') continue;
    var date=r[7] instanceof Date?r[7]:new Date((''+r[7]).split('T')[0]+'T12:00:00');
    var days=Math.floor((new Date(today.getFullYear(),today.getMonth(),today.getDate())-new Date(date.getFullYear(),date.getMonth(),date.getDate()))/86400000);
    var service=''+(r[5]||''), targetDays=service.toLowerCase().indexOf('descarga')>=0?2:1;
    if(days!==targetDays) continue;
    var first=(''+r[2]).trim().split(' ')[0];
    var msg='Hola '+first+', Â¿cÃ³mo te has sentido despuÃ©s de tu sesiÃ³n de '+service+'? Queremos acompaÃ±ar tu evoluciÃ³n. Si tienes alguna molestia o cambio, cuÃ©ntanos por aquÃ­.';
    if(queueAutomationMessage_('Seguimiento',r[2],r[3],msg,'Seguimiento '+targetDays+' dÃ­a(s) despuÃ©s',r[0],'followup|'+r[0])) queued++;
    if(cfg.autoFollowupEmail && r[4] && (''+r[4]).indexOf('@')>0) GmailApp.sendEmail(r[4],'Â¿CÃ³mo sigues despuÃ©s de tu sesiÃ³n?',msg,{name:'Cuidándote Fisioterapia'});
  }
  return {queued:queued};
}

function queueInactiveReminders_() {
  var data=getRemindersData();if(!data.ok)return{queued:0,error:data.error};
  var list=data.semana4.concat(data.semana5),queued=0,now=new Date();
  var weekKey=now.getFullYear()+'-'+Math.ceil((((now-new Date(now.getFullYear(),0,1))/86400000)+1)/7);
  for(var i=0;i<list.length;i++){
    var p=list[i],first=p.nombre.split(' ')[0];
    var msg='Hola '+first+', Â¿cÃ³mo te has sentido? Han pasado '+p.dias+' dÃ­as desde tu Ãºltima sesiÃ³n de '+p.lastServicio+'. Si quieres retomar tu proceso, tenemos horarios disponibles esta semana.';
    if(queueAutomationMessage_('ReactivaciÃ³n',p.nombre,p.telefono,msg,'Paciente sin regresar',p.nombre,'inactive|'+p.nombre.toLowerCase()+'|'+weekKey))queued++;
  }
  return{queued:queued};
}

function getPendingPayments_() {
  var rows=getOrCreateSheet().getSheetByName('Citas').getDataRange().getValues(), today=fmtDate(new Date()), out=[];
  for(var i=1;i<rows.length;i++) {
    var r=rows[i], date=r[7] instanceof Date?fmtDate(r[7]):(''+(r[7]||'')).split('T')[0];
    if(date<today && (''+r[10]).toLowerCase()!=='cancelada' && !(''+(r[14]||'')).trim()) out.push({id:r[0],nombre:r[2],fecha:date,precio:r[9]});
  }
  return out;
}

function sendPendingPaymentsSummary_() {
  var items=getPendingPayments_();
  if(!items.length) return {count:0};
  var body='Cobros pendientes detectados automÃ¡ticamente:\n\n'+items.slice(0,30).map(function(x){return 'â€¢ '+x.nombre+' Â· '+x.fecha+' Â· '+x.precio;}).join('\n')+'\n\nAbre el panel â†’ Centro de acciones para gestionarlos.';
  GmailApp.sendEmail(JESSICA_EMAIL,'Cobros pendientes â€” '+items.length,body);
  return {count:items.length};
}

function sendWeeklyManagementReport_() {
  var rows=getOrCreateSheet().getSheetByName('Citas').getDataRange().getValues(), now=new Date(), start=new Date(now);start.setDate(now.getDate()-7);
  var sessions=0,cancel=0,revenue=0,newPatients={};
  for(var i=1;i<rows.length;i++) {
    var r=rows[i],d=r[7] instanceof Date?r[7]:new Date((''+r[7]).split('T')[0]+'T12:00:00');
    if(d<start||d>now)continue;
    if((''+r[10]).toLowerCase()==='cancelada')cancel++;else{sessions++;newPatients[(''+r[2]).toLowerCase()]=1;revenue+=parseMoney_(r[9]);}
  }
  var pending=getPendingPayments_().length;
  var body='RESUMEN AUTOMÃTICO SEMANAL\n\nSesiones: '+sessions+'\nPacientes Ãºnicos: '+Object.keys(newPatients).length+'\nCancelaciones: '+cancel+'\nVentas registradas: $'+revenue.toLocaleString('es-CO')+'\nCobros pendientes: '+pending+'\n\nRevisa Indicadores de GestiÃ³n para tendencias y acciones.';
  GmailApp.sendEmail(JESSICA_EMAIL,'Resumen semanal de gestiÃ³n',body);
  return {sessions:sessions,cancelled:cancel,revenue:revenue,pending:pending};
}

function sendDataQualitySummary_() {
  var rows=getOrCreateSheet().getSheetByName('Pacientes').getDataRange().getValues(), missing=[];
  for(var i=1;i<rows.length;i++){var miss=[];if(!rows[i][1])miss.push('telÃ©fono');if(!rows[i][2])miss.push('email');if(miss.length)missing.push(rows[i][0]+' ('+miss.join(', ')+')');}
  if(missing.length) GmailApp.sendEmail(JESSICA_EMAIL,'Fichas de pacientes incompletas â€” '+missing.length,'Completar esta semana:\n\n'+missing.slice(0,50).join('\n'));
  return {incomplete:missing.length};
}

function parseMoney_(v) { return parseInt((''+(v||0)).replace(/\D/g,''),10)||0; }

function saveKPISnapshot_() {
  var ss=getOrCreateSheet(), rows=ss.getSheetByName('Citas').getDataRange().getValues(), now=new Date(), month=Utilities.formatDate(now,'America/Bogota','yyyy-MM');
  var sessions=0,cancelled=0,revenue=0,paid=0,patients={};
  for(var i=1;i<rows.length;i++){
    var r=rows[i],d=r[7] instanceof Date?fmtDate(r[7]):(''+(r[7]||'')).split('T')[0];if(d.slice(0,7)!==month)continue;
    if((''+r[10]).toLowerCase()==='cancelada')cancelled++;else{sessions++;revenue+=parseMoney_(r[9]);patients[(''+r[2]).toLowerCase()]=1;if((''+(r[14]||'')).trim())paid+=parseMoney_(r[9]);}
  }
  var survey={};try{survey=getEncuestaStats_();}catch(e){}
  var sh=getAutomationSheet_('KPIHistory',['month','updated','sessions','cancelled','revenue','paid','patients','nps','surveyResponses']);
  var data=sh.getDataRange().getValues(), row=0;for(var j=1;j<data.length;j++)if((''+data[j][0])===month){row=j+1;break;}
  var values=[month,new Date(),sessions,cancelled,revenue,paid,Object.keys(patients).length,survey.nps===undefined?'':survey.nps,survey.totalMes||0];
  if(row)sh.getRange(row,1,1,values.length).setValues([values]);else sh.appendRow(values);
  return {month:month,sessions:sessions,revenue:revenue};
}

function getKPIHistory_() {
  var sh=getAutomationSheet_('KPIHistory',['month','updated','sessions','cancelled','revenue','paid','patients','nps','surveyResponses']);
  var rows=sh.getDataRange().getValues(),items=[];
  for(var i=1;i<rows.length;i++)if(rows[i][0])items.push({month:''+rows[i][0],sessions:+rows[i][2]||0,cancelled:+rows[i][3]||0,revenue:+rows[i][4]||0,paid:+rows[i][5]||0,patients:+rows[i][6]||0,nps:rows[i][7]===''?null:+rows[i][7],surveyResponses:+rows[i][8]||0});
  return{ok:true,items:items.slice(-24)};
}

function createSpreadsheetBackup_() {
  var ss=getOrCreateSheet(), file=DriveApp.getFileById(ss.getId()), stamp=Utilities.formatDate(new Date(),'America/Bogota','yyyy-MM-dd_HHmm');
  var copy=file.makeCopy('BACKUP '+SS_NAME+' '+stamp);
  return {ok:true,name:copy.getName(),id:copy.getId()};
}

// -------------------------------------------------------------
// LISTA DE ESPERA SINCRONIZADA
// -------------------------------------------------------------
function waitlistSheet_(){return getAutomationSheet_('ListaEspera',['id','nombre','telefono','servicio','preferencia','estado','creado']);}
function getWaitlist(){var rows=waitlistSheet_().getDataRange().getValues(),items=[];for(var i=1;i<rows.length;i++)if(rows[i][0]&&(''+rows[i][5])!=='Retirado')items.push({id:''+rows[i][0],nombre:''+rows[i][1],telefono:''+rows[i][2],servicio:''+rows[i][3],preferencia:''+rows[i][4],estado:''+rows[i][5],creado:rows[i][6] instanceof Date?rows[i][6].toISOString():''+rows[i][6]});return{ok:true,items:items.reverse()};}
function addWaitlist(data){try{var p=JSON.parse(decodeURIComponent(data||'{}'));if(!p.nombre||!p.telefono)return{ok:false,error:'Nombre y telÃ©fono son obligatorios'};var id='WAIT-'+new Date().getTime();waitlistSheet_().appendRow([id,p.nombre,p.telefono,p.servicio||'',p.preferencia||'','Esperando',new Date()]);return{ok:true,id:id};}catch(e){return{ok:false,error:e.message};}}
function removeWaitlist(id){var sh=waitlistSheet_(),rows=sh.getDataRange().getValues();for(var i=1;i<rows.length;i++)if((''+rows[i][0])===(''+id)){sh.getRange(i+1,6).setValue('Retirado');return{ok:true};}return{ok:false,error:'Paciente no encontrado'};}

function queueWaitlistMatch_(booking) {
  if(!getAutomationConfig().waitlistMatching)return 0;
  var list=getWaitlist().items,queued=0;
  for(var i=0;i<list.length;i++){
    var p=list[i];if(p.servicio&&booking.servicio&&p.servicio.toLowerCase().indexOf((''+booking.servicio).toLowerCase())<0&&(''+booking.servicio).toLowerCase().indexOf(p.servicio.toLowerCase())<0)continue;
    var msg='Hola '+p.nombre.split(' ')[0]+', se liberÃ³ un horario el '+booking.fecha+' a las '+booking.hora+' para '+booking.servicio+'. Â¿Te gustarÃ­a tomarlo?';
    if(queueAutomationMessage_('Lista de espera',p.nombre,p.telefono,msg,'Horario liberado',booking.id,'wait|'+booking.id+'|'+p.id))queued++;
  }
  return queued;
}



