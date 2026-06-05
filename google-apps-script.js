// =============================================================
//  JESSICA OCAMPO FISIOTERAPEUTA — Apps Script Backend
//  Funciones: Reservas, Base de datos, Disponibilidad,
//             Panel Admin, Recordatorios diarios
// =============================================================

// IMPORTANTE: estas variables se leen desde PropertiesService (no están en código).
// Para configurarlas: en el editor de Apps Script → Proyecto → Propiedades del script → agrega:
//   ADMIN_TOKEN   → tu contraseña admin (ej: una cadena larga aleatoria)
//   GEMINI_API_KEY → tu clave de Gemini AI Studio
var _props        = PropertiesService.getScriptProperties();
var ADMIN_TOKEN   = _props.getProperty('ADMIN_TOKEN')   || 'JESSICA2026';
var GEMINI_API_KEY = _props.getProperty('GEMINI_API_KEY') || '';

// ── SESIONES ── token UUID almacenado en CacheService (TTL 4 horas)
function generateSessionToken() {
  var chars = 'abcdefghijklmnopqrstuvwxyz0123456789';
  var t = '';
  for (var i = 0; i < 48; i++) t += chars.charAt(Math.floor(Math.random() * chars.length));
  return t;
}
function createSession() {
  var token = generateSessionToken();
  CacheService.getScriptCache().put('sess_' + token, '1', 14400); // 4 horas
  return token;
}
function validateSession(token) {
  if (!token || token.length < 20) return false;
  return CacheService.getScriptCache().get('sess_' + token) === '1';
}

// ── RATE LIMITING LOGIN ── máx 5 intentos fallidos en 5 minutos (global)
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
//  GET  — Disponibilidad / Datos admin / Acciones admin
// -------------------------------------------------------------
function doGet(e) {
  var p = e.parameter;

  if (p.test) {
    return txt('OK - Calendario: ' + CalendarApp.getDefaultCalendar().getName());
  }

  if (p.action === 'availability' && p.date) {
    return js(getAvailability(p.date, p.service));
  }

  // Pasaporte — lectura pública (sin token)
  if (p.action === 'getPassport' && p.nombre) {
    return js(getPassport(decodeURIComponent(p.nombre)));
  }

  // Reseñas Google — público (sin token)
  if (p.action === 'getReviews') {
    return js(getGoogleReviews());
  }

  if (!validateSession(p.token)) {
    return js({ok: false, error: 'Sin permiso'});
  }
  // Ventana deslizante: renovar TTL en cada acción válida
  CacheService.getScriptCache().put('sess_' + p.token, '1', 14400);

  if (p.action === 'ping')          return js({ok: true});
  if (p.action === 'adminData')     return js(getAdminData());
  if (p.action === 'block')         return js(doBlock(p));
  if (p.action === 'unblock')       return js(doUnblock(p));
  if (p.action === 'updateStatus')  return js(doUpdateStatus(p));
  if (p.action === 'adminBook')     return js(createBooking(JSON.parse(p.data), true));
  if (p.action === 'getCalEvents')  return js(getCalendarEvents(p.from, p.to));
  if (p.action === 'cancelBooking') return js(doCancelBooking(p.id));
  if (p.action === 'editBooking')   return js(doEditBooking(JSON.parse(p.data)));
  if (p.action === 'deletePatient')  return js(deletePatient(decodeURIComponent(p.nombre)));
  if (p.action === 'editPatient')    return js(editPatient(JSON.parse(p.data)));
  if (p.action === 'getReminders')   return js(getRemindersData());
  if (p.action === 'sendReminders')  return js(sendEmailReminders());
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

  // Pasaporte — escritura (requiere token admin)
  if (p.action === 'savePassport' && p.nombre) {
    return js(savePassport(decodeURIComponent(p.nombre), p.passport || '{}', p.descarga || '{}'));
  }

  return txt('Jessica Ocampo Fisioterapeuta - Sistema activo');
}

// -------------------------------------------------------------
//  POST — Reservas de pacientes + Evaluación Express con fotos
// -------------------------------------------------------------
function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);
    if (d.action === 'adminLogin') {
      if (!loginAllowed()) return js({ok: false, error: 'Demasiados intentos fallidos. Espera 5 minutos.'});
      if (!d.password || d.password !== ADMIN_TOKEN) {
        recordLoginFail();
        return js({ok: false, error: 'Credenciales incorrectas'});
      }
      resetLoginFails();
      var sessionToken = createSession();
      var adminData = getAdminData();
      adminData.sessionToken = sessionToken;
      return js(adminData);
    }
    if (d.action === 'changePassword') {
      if (!validateSession(d.token)) return js({ok: false, error: 'Sin permiso'});
      if (!d.currentPassword || d.currentPassword !== ADMIN_TOKEN) return js({ok: false, error: 'La contraseña actual es incorrecta.'});
      if (!d.newPassword || d.newPassword.length < 8) return js({ok: false, error: 'La nueva contraseña debe tener al menos 8 caracteres.'});
      PropertiesService.getScriptProperties().setProperty('ADMIN_TOKEN', d.newPassword);
      return js({ok: true});
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
// Servicios que son solo registro de paciente — NO crean cita en Google Calendar
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

  // Lock para evitar duplicados por peticiones simultáneas (race condition)
  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch(e) { return {ok: false, error: 'Sistema ocupado, intenta de nuevo'}; }

  try {
  var price = d.modality === 'Presencial' ? d.priceP : d.priceD;

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
    location: d.modality === 'Domicilio' ? (d.address || 'Domicilio - Pereira / Dosquebradas') : 'Pereira, Colombia'
  });
  event.addEmailReminder(60);
  event.addPopupReminder(30);

  // Guardar en Google Sheets (solo citas reales)
  var id    = 'C' + new Date().getTime();
  var phoneClean = ('' + (d.phone||'')).replace(/\D/g,'');
  cSheet.appendRow([
    id,
    new Date().toLocaleString('es-CO'),
    d.name, phoneClean, d.email,
    d.service, d.modality,
    d.date, d.time, price,
    start < new Date() ? 'Atendida' : 'Confirmada',
    d.address || '', d.notes || '', d.notaAdmin || ''
  ]);
  // Forzar columna Telefono como texto para evitar #ERROR! en Sheets
  cSheet.getRange(cSheet.getLastRow(), 4).setNumberFormat('@').setValue(phoneClean);

  // Guardar/actualizar paciente en hoja Pacientes
  upsertPaciente(d.name, d.phone, d.email);

  // No enviar correos cuando la cita la crea el admin
  if (isAdmin) return {ok: true, id: id};

  // Para citas reales: enviar todos los correos y WhatsApp
  var tel  = (d.phone || '').replace(/\D/g,'');
  if (tel.length <= 10) tel = '57' + tel;
  var _waDias = ['domingo','lunes','martes','miércoles','jueves','viernes','sábado'];
  var _waMeses = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
  var _waDP = d.date.split('-');
  var _waFechaObj = new Date(+_waDP[0], +_waDP[1]-1, +_waDP[2]);
  var _waFecha = _waDias[_waFechaObj.getDay()] + ' ' + +_waDP[2] + ' de ' + _waMeses[+_waDP[1]-1];
  var waConfirm = '✅ Cita confirmada, ' + d.name.split(' ')[0] + '!\n\n' +
    '📌 ' + d.service + '\n' +
    '   ' + _waFecha + ' · ' + d.time + ' · ' + d.modality + '\n\n' +
    'Hasta pronto. Gracias por confiar en nuestros servicios.\n' +
    '— Jessica Ocampo Fisioterapeuta';
  var waLink = 'https://wa.me/' + tel + '?text=' + encodeURIComponent(waConfirm);

  GmailApp.sendEmail(
    JESSICA_EMAIL,
    'Nueva cita: ' + d.name + ' - ' + d.service + ' | ' + d.date,
    buildEmailJessica(d, price) + '\n\n>> Confirmar al paciente por WhatsApp (1 clic):\n' + waLink + '\n\nID cita: ' + id
  );

  if (d.email && d.email.indexOf('@') > 0) {
    GmailApp.sendEmail(
      d.email,
      '✅ Cita confirmada — Jessica Ocampo Fisioterapeuta',
      'Tu cita está confirmada. Si no puedes ver este correo, contáctanos al +57 313 646 7945.',
      {htmlBody: buildEmailCliente(d, price), name: 'Jessica Ocampo Fisioterapeuta'}
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

// -------------------------------------------------------------
//  DISPONIBILIDAD — lee Sheets + Calendario UNA sola vez
// -------------------------------------------------------------
function getAvailability(date, service) {
  var SLOTS = ['07:00','08:00','09:00','10:00','11:00','13:00','14:00','15:00','16:00','17:00','18:00','19:00'];
  var result = {};
  var newDur = getServiceDuration(service); // duración del servicio que quiere agendar

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

    // 1. Google Calendar (eventos personales solamente — los [CITA] ya están en Sheets)
    for (var k = 0; k < calEvents.length && ok; k++) {
      var ev = calEvents[k];
      if (ev.isAllDayEvent()) continue;
      if (ev.getTitle().indexOf('[CITA]') === 0) continue; // ya chequeados vía Sheets
      if (start < ev.getEndTime() && endNew > ev.getStartTime()) ok = false;
    }

    // 2. Citas en Sheets — usar duración real de cada cita existente
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

  return {ok: true, date: date, slots: result};
}

function checkAvailability(date, time, modality, service) {
  var start = parseDT(date, time);
  var mins  = getServiceDuration(service) + (modality === 'Domicilio' ? 30 : 0);
  var end   = new Date(start.getTime() + mins * 60000);

  // 1. Google Calendar — bloquear solo eventos personales (no [CITA])
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

  // 2. Citas existentes — usar duración real de cada servicio
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
    // Eliminar por ID único si existe (bloqueos nuevos)
    if (p.bid && rows[i][5] && rows[i][5] === p.bid) {
      sheet.deleteRow(i + 1);
      return {ok: true};
    }
    // Fallback para bloqueos viejos sin ID: comparar fecha + hora en múltiples formatos
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
        events.push({title: ev.getTitle(), fecha: fmtDate(ev.getStartTime()), hora: 'Todo el día', allDay: true});
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
  return {ok: true};
}

// Edita una cita existente en Sheets y actualiza el evento del Calendar
function doEditBooking(d) {
  var sheet = getOrCreateSheet().getSheetByName('Citas');
  var rows  = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] !== d.id) continue;
    var oldFecha = sd(rows[i][7]);
    var oldHora  = st(rows[i][8]);
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
          var ns      = parseDT(d.fecha || oldFecha, d.hora || oldHora);
          var newServ = d.servicio || rows[i][5];
          var newMod  = d.modalidad || rows[i][6];
          var newMins = getServiceDuration(newServ) + (newMod === 'Domicilio' ? 30 : 0);
          calEvs[k].setTime(ns, new Date(ns.getTime() + newMins * 60000));
          if (d.servicio) calEvs[k].setTitle('[CITA] ' + d.servicio + ' - ' + rows[i][2]);
          break;
        }
      }
    } catch(x) {}
    return {ok: true};
  }
  return {ok: false, error: 'Cita no encontrada'};
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

// Edita nombre, teléfono y email en Citas y en Pacientes
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
  // Si no existía en Pacientes (solo en Citas), crearlo
  if (!updatedPac && d.newNombre) {
    var today = new Date().toLocaleDateString('es-CO');
    pSheet.appendRow([d.newNombre, phone, d.email || '', today, today]);
    pSheet.getRange(pSheet.getLastRow(), 2).setNumberFormat('@').setValue(phone);
  }

  return {ok: true, updated: count};
}

function getAdminData() {
  var ss = getOrCreateSheet();

  var cRows = ss.getSheetByName('Citas').getDataRange().getValues();
  var citas = [];
  for (var i = 1; i < cRows.length; i++) {
    var r = cRows[i];
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
      pago: ('' + (r[14] || '')).trim()
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
//  HELPERS PLANES — detección y lógica de pagos
// -------------------------------------------------------------
function infoPlan(serv, mod) {
  var s = (serv || '').split('(')[0].trim();
  var esDom = mod === 'Domicilio';
  var planes = {
    'Paquete Readaptación Inicio': { total:6,  pagoDosEn:4, mitadP:'$189.000', mitadD:'$234.500' },
    'Paquete Readaptación Avance': { total:8,  pagoDosEn:5, mitadP:'$238.000', mitadD:'$299.000' },
    'Paquete Readaptación Total':  { total:10, pagoDosEn:6, mitadP:'$280.000', mitadD:'$361.000' },
    'Paquete Recuperación Full':   { total:3,  pagoDosEn:null, mitadP:null, mitadD:null },
    'Combo Diagnóstico Pro':       { total:2,  pagoDosEn:null, mitadP:null, mitadD:null },
    'Combo Bienvenida':            { total:2,  pagoDosEn:null, mitadP:null, mitadD:null },
    'Plan Activo':                 { total:2,  pagoDosEn:null, mitadP:null, mitadD:null },
    'Plan Pro':                    { total:3,  pagoDosEn:null, mitadP:null, mitadD:null },
    'Mini-sesión Familiar 20 min': { total:1,  pagoDosEn:null, mitadP:null, mitadD:null },
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
  var cuando = esHoy ? 'Hoy ' + fechaLegible : 'Mañana ' + fechaLegible;
  var encabezado = esHoy ? '🩺 *¡Hoy tienes cita!*' : '🩺 *Recordatorio de cita*';
  var progLine = (plan && sesionActual)
    ? '\n🔄 Sesión ' + sesionActual + (plan.total ? ' de ' + plan.total : '')
    : '';

  var pagoLine = '';
  if (esHoy) {
    if (plan && sesionActual) {
      if (sesionActual === 1) {
        pagoLine = '\n\n💳 Pago inicial' + (plan.mitad ? ': ' + plan.mitad : '') +
          '\nBancolombia Ahorros: 91257857099\nLlave: 1010124692\nNequi: 3136467945\nTitular: Jessica Andrea Ocampo Barbosa';
      } else if (plan.pagoDosEn && sesionActual === plan.pagoDosEn) {
        pagoLine = '\n\n💳 Segundo pago del plan' + (plan.mitad ? ': ' + plan.mitad : '') +
          '\nBancolombia Ahorros: 91257857099\nLlave: 1010124692\nNequi: 3136467945\nTitular: Jessica Andrea Ocampo Barbosa';
      }
    }
  }

  return encabezado + '\n━━━━━━━━━━━━━━━━━━━━\n' +
    '💆 ' + serv + '\n' +
    '📅 ' + cuando + '\n' +
    '🕘 ' + hora + ' · ' + mod +
    progLine +
    pagoLine + '\n\n' +
    '¿Tienes algún cambio? Escríbeme 🙏\n' +
    '_Jessica Ocampo Fisioterapeuta_';
}

// -------------------------------------------------------------
//  RECORDATORIOS DIARIOS — ejecutar con trigger 7am
// -------------------------------------------------------------
function sendReminders() {
  var diasSemana = ['domingo','lunes','martes','miércoles','jueves','viernes','sábado'];
  var meses = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];

  var ss   = getOrCreateSheet();
  var rows = ss.getSheetByName('Citas').getDataRange().getValues();

  var today    = fmtDate(new Date());
  var tmrwDate = new Date(); tmrwDate.setDate(tmrwDate.getDate() + 1);
  var tomorrow = fmtDate(tmrwDate);

  var linksHoy = [], linksMañana = [];

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

    // Detectar si es plan y calcular sesión actual
    var plan = infoPlan(serv, mod);
    var sesionActual = null;
    if (plan) {
      sesionActual = contarSesiones(rows, nombre, serv, fecha) + 1;
    }

    if (fecha === tomorrow) {
      if (email && email.indexOf('@') > 0) {
        var dp = fecha.split('-');
        var fObj = new Date(+dp[0], +dp[1]-1, +dp[2]);
        var fechaLegible = diasSemana[fObj.getDay()] + ' ' + +dp[2] + ' de ' + meses[+dp[1]-1];
        GmailApp.sendEmail(
          email,
          'Recordatorio: mañana tienes cita — Jessica Ocampo Fisioterapeuta',
          'Este correo requiere un cliente de correo con soporte HTML.',
          {htmlBody: buildReminderEmail(nombre, serv, fechaLegible, hora, mod, precio, false, plan, sesionActual),
           name: 'Jessica Ocampo Fisioterapeuta'}
        );
      }
      var msg1 = mensajePlanWA(nombre, serv, hora, mod, fechaLegible, plan, sesionActual, false);
      linksMañana.push(nombre + ' (' + hora + '): https://wa.me/' + phone + '?text=' + encodeURIComponent(msg1));
    }

    if (fecha === today) {
      if (email && email.indexOf('@') > 0) {
        var dp2 = fecha.split('-');
        var fObj2 = new Date(+dp2[0], +dp2[1]-1, +dp2[2]);
        var fechaLegible2 = diasSemana[fObj2.getDay()] + ' ' + +dp2[2] + ' de ' + meses[+dp2[1]-1];
        GmailApp.sendEmail(
          email,
          '⏰ Hoy tienes cita — Jessica Ocampo Fisioterapeuta',
          'Este correo requiere un cliente de correo con soporte HTML.',
          {htmlBody: buildReminderEmail(nombre, serv, fechaLegible2, hora, mod, precio, true, plan, sesionActual),
           name: 'Jessica Ocampo Fisioterapeuta'}
        );
      }
      var msg2 = mensajePlanWA(nombre, serv, hora, mod, fechaLegible2, plan, sesionActual, true);
      linksHoy.push(nombre + ' (' + hora + '): https://wa.me/' + phone + '?text=' + encodeURIComponent(msg2));
    }
  }

  // Resumen diario para Jessica con links de WhatsApp 1-clic
  if (linksHoy.length > 0 || linksMañana.length > 0) {
    var body = 'Recordatorios automáticos del día ' + today + '\n\n';
    if (linksHoy.length)    body += '== CITAS DE HOY (WhatsApp 1 clic) ==\n' + linksHoy.join('\n') + '\n\n';
    if (linksMañana.length) body += '== CITAS DE MAÑANA (WhatsApp 1 clic) ==\n' + linksMañana.join('\n') + '\n';
    GmailApp.sendEmail(JESSICA_EMAIL, 'Resumen de citas - ' + today, body);
  }
}

// Ejecuta ESTA funcion UNA sola vez para activar los recordatorios diarios:
function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('sendReminders').timeBased().everyDays(1).atHour(7).inTimezone('America/Bogota').create();
  ScriptApp.newTrigger('autoMarcarAtendidas').timeBased().everyDays(1).atHour(22).inTimezone('America/Bogota').create();
  Logger.log('Triggers activados: sendReminders 7am y autoMarcarAtendidas 10pm hora Colombia.');
}

// -------------------------------------------------------------
//  AUTO MARCAR ATENDIDAS — trigger diario a las 10pm
// -------------------------------------------------------------
function autoMarcarAtendidas() {
  var ss    = getOrCreateSheet();
  var sheet = ss.getSheetByName('Citas');
  var rows  = sheet.getDataRange().getValues();
  var ahora = new Date();
  var hoy   = fmtDate(ahora);
  var count = 0;

  for (var i = 1; i < rows.length; i++) {
    var r      = rows[i];
    var estado = ('' + (r[10] || '')).trim();
    // Solo aplica a citas Confirmadas o Pendientes — no tocar Canceladas, Atendidas, No asistió
    if (estado !== 'Confirmada' && estado !== 'Pendiente') continue;

    var fecha = (r[7] instanceof Date) ? fmtDate(r[7]) : ('' + r[7]).split('T')[0];
    var hora  = st(r[8]);

    // Citas de días anteriores → marcar como Atendida directamente
    if (fecha < hoy) {
      sheet.getRange(i + 1, 11).setValue('Atendida');
      count++;
      continue;
    }

    // Citas de hoy → marcar solo si la hora ya pasó (+ 30 min de margen)
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
    // Crear hoja Pacientes si no existe aún
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
        // Actualizar teléfono/email si llegaron nuevos y actualizar última visita
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
  var s = (service || '').toLowerCase()
    .replace(/[áàâ]/g,'a').replace(/[éèê]/g,'e')
    .replace(/[íìî]/g,'i').replace(/[óòô]/g,'o').replace(/[úùû]/g,'u');
  if (s.indexOf('mini') > -1)           return 20;  // Mini-sesión Familiar 20 min
  if (s.indexOf('completa') > -1)       return 80;  // Descarga Muscular Completa
  if (s.indexOf('full') > -1)           return 80;  // Paquete Recuperación Full
  if (s.indexOf('plan pro') > -1)       return 80;  // Plan Pro (sesión Full incluida)
  if (s.indexOf('paquete readap') > -1) return 45;  // Paquetes Readaptación Inicio/Avance/Total
  if (s.indexOf('readaptacion') > -1)   return 50;  // Readaptación Funcional suelta
  if (s.indexOf('valoracion') > -1)     return 50;  // Valoración Funcional
  if (s.indexOf('piernas') > -1)        return 50;  // Descarga Muscular Piernas
  if (s.indexOf('cuello') > -1)         return 50;  // Descarga Muscular Cuello/Espalda
  if (s.indexOf('plan activo') > -1)    return 50;  // Plan Activo (Express)
  if (s.indexOf('combo') > -1)          return 50;  // Combo Diagnóstico Pro / Combo Bienvenida
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
  var diasSemana = ['domingo','lunes','martes','miércoles','jueves','viernes','sábado'];
  var meses = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
  var dp = d.date.split('-');
  var fechaObj = new Date(+dp[0], +dp[1]-1, +dp[2]);
  var fechaLegible = diasSemana[fechaObj.getDay()] + ' ' + +dp[2] + ' de ' + meses[+dp[1]-1] + ' de ' + dp[0];
  var modDetalle = d.modality + (d.address ? ' — ' + d.address : '');
  var primerNombre = d.name.split(' ')[0];

  return '<div style="font-family:Arial,sans-serif;max-width:520px;margin:0 auto;border:1px solid #e5e7eb;border-radius:12px;overflow:hidden">' +
    '<div style="background:#0d9488;padding:20px 32px;text-align:center">' +
    '<p style="color:#fff;margin:0;font-size:15px;font-weight:600">🩺 Jessica Ocampo Fisioterapeuta</p>' +
    '</div>' +
    '<div style="padding:28px 32px">' +
    '<p style="font-size:17px;font-weight:700;margin:0 0 20px;color:#111827">✅ Cita confirmada, ' + primerNombre + '!</p>' +
    '<div style="margin:0 0 20px">' +
    '<p style="margin:0 0 6px;font-size:14px;font-weight:600;color:#111827">📌 ' + d.service + '</p>' +
    '<p style="margin:0;font-size:13px;color:#6b7280">' + fechaLegible + ' · ' + d.time + ' · ' + modDetalle + '</p>' +
    '</div>' +
    '<hr style="border:none;border-top:2px solid #e5e7eb;margin:20px 0">' +
    '<p style="font-size:13px;color:#6b7280;margin:0 0 16px">📩 Recibirás un recordatorio el día anterior y el mismo día de tu cita.</p>' +
    '<hr style="border:none;border-top:1px solid #e5e7eb;margin:20px 0">' +
    '<p style="font-size:14px;color:#374151;margin:0 0 4px">Hasta pronto. Gracias por confiar en nuestros servicios.</p>' +
    '<p style="font-size:14px;color:#374151;margin:0;font-style:italic">— Jessica Ocampo Fisioterapeuta</p>' +
    '</div>' +
    '<div style="background:#f9fafb;padding:14px 32px;text-align:center;font-size:12px;color:#9ca3af">' +
    'Jessica Ocampo Fisioterapeuta · Pereira, Colombia · ' +
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
    ? '¡Hola ' + primerNombre + '! 🌟 Hoy es el día de tu sesión de <strong>' + serv + '</strong> a las <strong>' + hora + '</strong> (' + mod + ').'
    : '¡Hola ' + primerNombre + '! 👋 Te recuerdo que mañana tienes tu sesión de <strong>' + serv + '</strong> a las <strong>' + hora + '</strong> (' + mod + ').';

  var progStr = '';
  if (plan && sesionActual) {
    progStr = '<p style="margin:10px 0 0;font-size:13px;color:#6b7280">🔄 ' +
      (plan.total ? 'Sesión ' + sesionActual + ' de ' + plan.total : 'Sesión ' + sesionActual) + '</p>';
  }

  // Bloque de pago: solo en recordatorio del mismo día
  var bloquePago = '';
  if (esHoy) {
    if (plan && sesionActual) {
      if (sesionActual === 1) {
        bloquePago = '<hr style="border:none;border-top:1px solid #e5e7eb;margin:20px 0">' +
          buildPaymentBlock('💰 Pago inicial del plan' + (plan.mitad ? ' — ' + plan.mitad : ''), 'Para comenzar recuerda traer el pago inicial.');
      } else if (plan.pagoDosEn && sesionActual === plan.pagoDosEn) {
        bloquePago = '<hr style="border:none;border-top:1px solid #e5e7eb;margin:20px 0">' +
          buildPaymentBlock('💰 Segundo pago del plan' + (plan.mitad ? ' — ' + plan.mitad : ''), 'Esta sesión corresponde al segundo y último pago de tu plan. Recuerda traerlo.');
      }
    } else if (!plan && precio) {
      bloquePago = '<hr style="border:none;border-top:1px solid #e5e7eb;margin:20px 0">' +
        buildPaymentBlock('💰 Valor: ' + precio, null);
    }
  }

  return '<div style="font-family:Arial,sans-serif;max-width:520px;margin:0 auto;border:1px solid #e5e7eb;border-radius:12px;overflow:hidden">' +
    '<div style="background:' + (esHoy ? '#0284c7' : '#0d9488') + ';padding:20px 32px;text-align:center">' +
    '<p style="color:#fff;margin:0;font-size:15px;font-weight:600">🩺 Jessica Ocampo Fisioterapeuta</p>' +
    '</div>' +
    '<div style="padding:28px 32px">' +
    '<p style="font-size:14px;line-height:1.7;color:#111827;margin:0">' + intro + '</p>' +
    progStr +
    bloquePago +
    '<hr style="border:none;border-top:1px solid #e5e7eb;margin:20px 0">' +
    '<p style="font-size:13px;color:#6b7280;margin:0 0 16px">¿Tienes algún cambio? <a href="https://wa.me/573136467945" style="color:#0d9488">Escríbeme</a>.</p>' +
    '<p style="font-size:14px;color:#374151;margin:0 0 4px">Hasta pronto. Gracias por confiar en nuestros servicios.</p>' +
    '<p style="font-size:14px;color:#374151;margin:0;font-style:italic">— Jessica Ocampo Fisioterapeuta</p>' +
    '</div>' +
    '<div style="background:#f9fafb;padding:14px 32px;text-align:center;font-size:12px;color:#9ca3af">' +
    'Jessica Ocampo Fisioterapeuta · Pereira, Colombia · ' +
    '<a href="https://wa.me/573136467945" style="color:#0d9488">+57 313 646 7945</a>' +
    '</div></div>';
}

// =============================================================
//  ADMIN KV — almacenamiento clave-valor sincronizado entre dispositivos
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

    // Construir índice key → número de fila (1-based)
    var keyToRow = {};
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) keyToRow['' + data[i][0]] = i + 1;
    }

    for (var key in updates) {
      var val = updates[key];
      if (val === '__DELETE__') {
        if (keyToRow[key]) {
          sh.deleteRow(keyToRow[key]);
          // Reconstruir índice tras borrar
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
    sh.appendRow(['nombre', 'passport', 'descarga', 'actualizado']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function getPassport(nombre) {
  try {
    var sh   = getPasaportesSheet();
    var data = sh.getDataRange().getValues();
    var norm = (nombre || '').toLowerCase().trim();
    for (var i = 1; i < data.length; i++) {
      if ((data[i][0] || '').toLowerCase().trim() === norm) {
        return {
          ok:       true,
          passport: data[i][1] ? JSON.parse(data[i][1]) : {},
          descarga: data[i][2] ? JSON.parse(data[i][2]) : {}
        };
      }
    }
    return { ok: true, passport: {}, descarga: {} };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

function savePassport(nombre, passportJson, descargaJson) {
  try {
    var sh   = getPasaportesSheet();
    var data = sh.getDataRange().getValues();
    var norm = (nombre || '').toLowerCase().trim();
    var now  = new Date().toISOString();
    for (var i = 1; i < data.length; i++) {
      if ((data[i][0] || '').toLowerCase().trim() === norm) {
        sh.getRange(i + 1, 1, 1, 4).setValues([[nombre, passportJson, descargaJson, now]]);
        return { ok: true };
      }
    }
    sh.appendRow([nombre, passportJson, descargaJson, now]);
    return { ok: true };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// -------------------------------------------------------------
//  RECORDATORIOS MENSUALES DE REAGENDAMIENTO
// -------------------------------------------------------------

// Devuelve pacientes cuya última cita fue hace ~4 semanas (semana4) o 5+ semanas (semana5)
function getRemindersData() {
  var ss   = getOrCreateSheet();
  var rows = ss.getSheetByName('Citas').getDataRange().getValues();

  // Construir mapa: último registro por paciente
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

// Envía emails a todos los pacientes con email registrado que están en semana 4 o 5+
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
      ? ('⏰ ' + primero + ', ya es momento de tu próxima descarga muscular')
      : ('💆 ' + primero + ', lleva 5 semanas desde tu última sesión');

    try {
      GmailApp.sendEmail(p.email, asunto, '', {
        htmlBody: buildReminderMensualEmail(p.nombre, item.semanas),
        name: 'Jessica Ocampo Fisioterapeuta'
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
      'Recordatorios de reagendamiento enviados — ' + sent + ' email(s)',
      'Resumen del envío automático de recordatorios:\n\n' +
      '✅ Emails enviados: ' + sent + '\n' +
      '⏭ Sin email (WhatsApp manual): ' + skipped + '\n' +
      '❌ Errores: ' + errors + '\n\n' +
      'Entra al panel admin → Recordatorios para enviarles WhatsApp a los pacientes sin email.');
  }

  return { ok: true, sent: sent, skipped: skipped, errors: errors };
}

function buildReminderMensualEmail(nombre, semanas) {
  var primero = nombre.split(' ')[0];
  var msg = semanas === 4
    ? ('Ya vamos en la <strong>semana 4</strong> desde tu última descarga muscular — la próxima semana sería el momento ideal para hacerla antes de que el cuerpo empiece a acumular tensión de nuevo.')
    : ('Ya se cumplieron las <strong>5 semanas</strong> desde tu última sesión de descarga — es el momento de reagendar. Mantener la frecuencia es lo que hace que los resultados se sostengan.');

  return '<div style="font-family:Arial,sans-serif;max-width:520px;margin:0 auto;border:1px solid #e5e7eb;border-radius:12px;overflow:hidden">' +
    '<div style="background:#0d9488;padding:24px 32px;text-align:center">' +
    '<h1 style="color:#fff;margin:0;font-size:19px">⏰ Es momento de tu próxima sesión</h1>' +
    '<p style="color:#ccfbf1;margin:6px 0 0;font-size:13px">Jessica Ocampo Fisioterapeuta</p>' +
    '</div>' +
    '<div style="padding:28px 32px">' +
    '<p style="margin:0 0 14px;font-size:15px">Hola <strong>' + primero + '</strong>! 👋 Soy Jessica Ocampo Fisioterapeuta.</p>' +
    '<p style="margin:0 0 24px;font-size:15px;line-height:1.65;color:#374151">' + msg + '</p>' +
    '<div style="text-align:center;margin:24px 0">' +
    '<a href="https://jessicaocampoft-ctrl.github.io/#agenda" style="background:#0d9488;color:#fff;padding:14px 32px;border-radius:8px;text-decoration:none;font-weight:600;font-size:15px;display:inline-block">Agendar mi cita 📅</a>' +
    '</div>' +
    '<hr style="border:none;border-top:1px solid #e5e7eb;margin:24px 0">' +
    '<p style="font-size:13px;color:#6b7280;margin:0 0 12px">¿Cómo ha sido tu experiencia? Tu opinión me ayuda a mejorar:</p>' +
    '<div style="text-align:center;margin:0 0 20px">' +
    '<a href="https://forms.gle/srX1enyKN59n8TfQA" style="background:#f9fafb;border:1px solid #e5e7eb;color:#0d9488;padding:11px 24px;border-radius:8px;text-decoration:none;font-weight:600;font-size:14px;display:inline-block">⭐ Responder encuesta de satisfacción</a>' +
    '</div>' +
    '<p style="font-size:13px;color:#6b7280;margin:0">¿Prefieres escribirme directamente?<br>' +
    '<a href="https://wa.me/573136467945" style="color:#0d9488">+57 313 646 7945 (WhatsApp)</a></p>' +
    '</div>' +
    '<div style="background:#f9fafb;padding:16px 32px;text-align:center;font-size:12px;color:#9ca3af">' +
    'Jessica Ocampo Fisioterapeuta · Pereira, Colombia<br>' +
    '<a href="https://jessicaocampoft-ctrl.github.io" style="color:#0d9488">jessicaocampoft-ctrl.github.io</a>' +
    '</div></div>';
}

// Ejecuta ESTA función UNA sola vez para activar el trigger semanal automático:
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

// ── RESEÑAS GOOGLE ──
function getGoogleReviews() {
  var PLACE_ID = 'ChIJVwU1iJ15sCARAQ_jFCdVsXI';
  var API_KEY  = 'AIzaSyAKtsK8EaAG0GE_0Ma-mNoaMwy1ZG0gEv8';
  try {
    var url = 'https://places.googleapis.com/v1/places/' + PLACE_ID + '?key=' + API_KEY;
    var res  = UrlFetchApp.fetch(url, {
      headers: { 'X-Goog-FieldMask': 'rating,userRatingCount,reviews.rating,reviews.text,reviews.originalText,reviews.authorAttribution,reviews.relativePublishTimeDescription' },
      muteHttpExceptions: true
    });
    var data = JSON.parse(res.getContentText());
    if (data.error) return { ok: false, error: data.error.message };
    return { ok: true, data: data };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// =============================================================
//  EVALUACIÓN EXPRESS CROSSFIT — Generación de reporte con IA
// =============================================================

function generateEvalReport(d, photos) {
  photos = photos || {};
  try {
    if (!GEMINI_API_KEY || GEMINI_API_KEY === 'PEGA_AQUI_TU_CLAVE_GEMINI') {
      return { ok: false, error: 'Configura GEMINI_API_KEY en el script. Obtenla gratis en aistudio.google.com' };
    }

    var parts = [{ text: buildEvalPrompt(d) }];

    // Añadir fotos posturales y de tests a la solicitud multimodal
    var photoLabels = {
      frontal:  'FOTO POSTURAL FRONTAL — analiza alineación de cabeza, hombros, crestas ilíacas, rodillas y pies',
      lateral:  'FOTO POSTURAL LATERAL — analiza adelantamiento cefálico, cifosis, lordosis, posición rodilla',
      posterior:'FOTO POSTURAL POSTERIOR — analiza escoliosis, asimetría escapular, pies y talones',
      ds:       'FOTO DEEP SQUAT — analiza profundidad, posición de rodillas, talones y tronco',
      oh:       'FOTO OVERHEAD REACH — analiza contacto de manos con pared y compensaciones',
      slsd:     'FOTO SINGLE LEG SQUAT DERECHO — analiza control de rodilla y cadera',
      slsi:     'FOTO SINGLE LEG SQUAT IZQUIERDO — analiza control de rodilla y cadera',
      shd:      'FOTO SHOULDER CLEARING MANO D ARRIBA — analiza distancia entre puños',
      shi:      'FOTO SHOULDER CLEARING MANO I ARRIBA — analiza distancia entre puños',
      bmd:      'FOTO BALANCE MONOPODAL DERECHO — analiza postura y estrategia de equilibrio',
      bmi:      'FOTO BALANCE MONOPODAL IZQUIERDO — analiza postura y estrategia de equilibrio',
      trdd:     'FOTO TRENDELENBURG APOYO DERECHO — analiza nivel pélvico y caída contralateral',
      trdi:     'FOTO TRENDELENBURG APOYO IZQUIERDO — analiza nivel pélvico y caída contralateral',
      sbd:      'FOTO SHIN BOX PIERNA DERECHA — analiza posición de shin y rango de movimiento',
      sbi:      'FOTO SHIN BOX PIERNA IZQUIERDA — analiza posición de shin y rango de movimiento',
      bdd:      'FOTO BIRD DOG LADO DERECHO — analiza neutralidad lumbar y control rotacional',
      bdi:      'FOTO BIRD DOG LADO IZQUIERDO — analiza neutralidad lumbar y control rotacional',
      dbd:      'FOTO DEAD BUG LADO DERECHO — analiza neutro lumbar al extender extremidades',
      dbi:      'FOTO DEAD BUG LADO IZQUIERDO — analiza neutro lumbar al extender extremidades',
      pk:       'FOTO PLANK — analiza alineación de cadera, espalda y posición general'
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
      parts.push({ text: '\nCon base en las ' + photoCount + ' fotografías anteriores y los datos ingresados, incluye en tu análisis observaciones específicas de lo que ves en las imágenes de cada test. Valida o corrige los scores asignados por la fisioterapeuta si lo consideras necesario, explicando el motivo.' });
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
    sh.appendRow(['ID','Título','Tipo','Fecha','HoraInicio','HoraFin','Duración','Cobro','Notas']);
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
  // Forzar fecha y horas como texto para evitar auto-detección de Sheets
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
//  SISTEMA DE CÓDIGOS — REF-MES-NNN  /  BONO-MES-NNN
// =============================================================

function getCodigosSheet() {
  var ss = getOrCreateSheet();
  var sh = ss.getSheetByName('Codigos');
  if (!sh) {
    sh = ss.insertSheet('Codigos');
    sh.appendRow(['Código','Tipo','Paciente','Teléfono','Referido por','Fecha','Estado','CódigoRef']);
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
  // Formatear teléfono como texto
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
  return {ok: false, error: 'Código no encontrado'};
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

function buildEvalPrompt(d) {
  var hallazgos = (d.hallazgos && d.hallazgos.length)
    ? d.hallazgos.join(', ')
    : 'Sin hallazgos registrados';

  var screensText = '';
  if (d.screens) {
    var sc = d.screens;
    screensText += '\nRESULTADOS PRUEBAS FUNCIONALES (escala 0-3, donde 3=normal, 0=dolor/no puede):\n';
    if (sc.deepSquat) {
      screensText += '• Deep Squat: ' + (sc.deepSquat.score !== null ? sc.deepSquat.score + '/3' : 'No evaluado');
      if (sc.deepSquat.obs) screensText += ' — ' + sc.deepSquat.obs;
      screensText += '\n';
    }
    if (sc.overheadReach) {
      screensText += '• Overhead Reach: ' + (sc.overheadReach.score !== null ? sc.overheadReach.score + '/3' : 'No evaluado');
      if (sc.overheadReach.obs) screensText += ' — ' + sc.overheadReach.obs;
      screensText += '\n';
    }
    if (sc.singleLegSquat) {
      screensText += '• Single Leg Squat: D=' + (sc.singleLegSquat.scoreD !== null ? sc.singleLegSquat.scoreD + '/3' : 'No eval') +
        ', I=' + (sc.singleLegSquat.scoreI !== null ? sc.singleLegSquat.scoreI + '/3' : 'No eval');
      if (sc.singleLegSquat.obs) screensText += ' — ' + sc.singleLegSquat.obs;
      screensText += '\n';
    }
    if (sc.shoulderMob) {
      screensText += '• Shoulder Clearing: D=' + (sc.shoulderMob.scoreD !== null ? sc.shoulderMob.scoreD + '/3' : 'No eval') +
        ', I=' + (sc.shoulderMob.scoreI !== null ? sc.shoulderMob.scoreI + '/3' : 'No eval');
      if (sc.shoulderMob.obs) screensText += ' — ' + sc.shoulderMob.obs;
      screensText += '\n';
    }
    if (sc.trendelenburg) {
      screensText += '• Trendelenburg: D=' + (sc.trendelenburg.resultD || 'No eval') +
        ', I=' + (sc.trendelenburg.resultI || 'No eval');
      if (sc.trendelenburg.obs) screensText += ' — ' + sc.trendelenburg.obs;
      screensText += '\n';
    }
    if (sc.shinBox) {
      screensText += '• Shin Box: D RI=' + (sc.shinBox.riD || '?') + '° RE=' + (sc.shinBox.reD || '?') +
        '° Trans=' + (sc.shinBox.transD !== null && sc.shinBox.transD !== undefined ? sc.shinBox.transD + '/3' : 'No eval') +
        ' | I RI=' + (sc.shinBox.riI || '?') + '° RE=' + (sc.shinBox.reI || '?') +
        '° Trans=' + (sc.shinBox.transI !== null && sc.shinBox.transI !== undefined ? sc.shinBox.transI + '/3' : 'No eval');
      if (sc.shinBox.obs) screensText += ' — ' + sc.shinBox.obs;
      screensText += ' (Normal RI 35-45°, RE 40-60°)\n';
    }
    if (sc.birdDog) {
      screensText += '• Bird Dog (control core): D=' + (sc.birdDog.scoreD !== null && sc.birdDog.scoreD !== undefined ? sc.birdDog.scoreD + '/3' : 'No eval') +
        ', I=' + (sc.birdDog.scoreI !== null && sc.birdDog.scoreI !== undefined ? sc.birdDog.scoreI + '/3' : 'No eval');
      if (sc.birdDog.obs) screensText += ' — ' + sc.birdDog.obs;
      screensText += '\n';
    }
  }

  return 'Eres una fisioterapeuta deportiva especializada en CrossFit llamada Jessica Ocampo, con sede en Pereira, Colombia.\n' +
    'Genera un reporte de evaluación postural express profesional, empático y motivador para:\n\n' +
    'ATLETA: ' + d.nombre + (d.edad && d.edad !== 'N/A' ? ', ' + d.edad + ' años' : '') + '\n' +
    'NIVEL CROSSFIT: ' + (d.nivel || 'No especificado') + '\n' +
    'TIEMPO EN CROSSFIT: ' + (d.anios || 'No especificado') + '\n' +
    'OBJETIVO: ' + (d.objetivo || 'No especificado') + '\n' +
    'MOLESTIAS ACTUALES: ' + (d.molestias || 'Ninguna') + '\n' +
    'SEVERIDAD ESTIMADA: ' + (d.severidad || 'No especificada') + '\n' +
    'HALLAZGOS POSTURALES: ' + hallazgos + '\n' +
    screensText +
    (d.observaciones ? 'OBSERVACIONES ADICIONALES: ' + d.observaciones + '\n' : '') +
    '\nGenera DOS reportes separados. El primero es técnico (para la fisioterapeuta), el segundo es para el paciente (lenguaje simple). Sepáralos exactamente con la línea: ===PACIENTE===\n\n' +

    '━━━ REPORTE TÉCNICO (para la fisioterapeuta) ━━━\n' +
    'Usa terminología clínica. Secciones con títulos en negrilla:\n\n' +
    '**RESUMEN EJECUTIVO**\n' +
    '[2-3 oraciones: estado postural y funcional general, impacto en rendimiento CrossFit. Si hay fotos, describe hallazgos visuales relevantes.]\n\n' +
    '**ANÁLISIS VISUAL DE FOTOS**\n' +
    '[SOLO si se enviaron fotos: hallazgos por foto. Si no hay fotos, omite esta sección.]\n\n' +
    '**ANÁLISIS POR ZONAS**\n' +
    '[Por cada zona con hallazgos: significado biomecánico y afectación en movimientos CrossFit (snatch, clean, squat, deadlift). Omite zonas sin hallazgos.]\n\n' +
    '**ANÁLISIS FUNCIONAL**\n' +
    '[Interpreta Trendelenburg (glúteo medio), Shin Box (movilidad cadera), Bird Dog (control core). Correlaciona con hallazgos posturales.]\n\n' +
    '**RIESGOS IDENTIFICADOS**\n' +
    '[Máximo 4 riesgos concretos de lesión. Específicos para CrossFit. Viñetas con •]\n\n' +
    '**PLAN DE ACCIÓN RECOMENDADO**\n' +
    '[3-5 recomendaciones clínicas por prioridad: tipo de intervención en fisioterapia (movilidad, estabilización, corrección postural). Viñetas con •]\n\n' +

    '===PACIENTE===\n\n' +

    '━━━ REPORTE PARA EL PACIENTE ━━━\n' +
    'Lenguaje simple, cálido y motivador. Sin términos clínicos. Secciones:\n\n' +
    '**¿QUÉ ENCONTRAMOS HOY?**\n' +
    '[2-3 oraciones explicando en palabras simples lo que se encontró y cómo afecta su entrenamiento. Sin jerga médica.]\n\n' +
    '**LO QUE ESTÁ BIEN 💪**\n' +
    '[1-2 oraciones destacando los puntos fuertes del atleta. Siempre hay algo positivo.]\n\n' +
    '**RECOMENDACIONES PARA TU ENTRENAMIENTO**\n' +
    '[1 sola recomendación general, en 2-3 oraciones. Sin mencionar ejercicios específicos ni movimientos de CrossFit por nombre. Enfocada en un principio general (ej: técnica sobre velocidad, escuchar el cuerpo, etc.). Lenguaje simple.]\n\n' +
    '**TU SIGUIENTE PASO**\n' +
    '[1 párrafo corto y motivador invitándolo a agendar su plan de fisioterapia. Menciona que ya identificaste exactamente qué trabajar y que los resultados se ven rápido con un plan personalizado. Cálido y sin presión.]\n\n' +

    'REGLAS GLOBALES: Reporte técnico máx 500 palabras. Reporte paciente máx 300 palabras. Tono técnico: preciso y profesional. Tono paciente: cercano, claro, motivador.';
}

// =============================================================
//  ENCUESTA DE SATISFACCIÓN — NPS y % respuestas desde Google Forms
// =============================================================
// Ejecuta esta función UNA vez para autorizar el permiso de Formularios
function autorizarFormularios() {
  FormApp.openById('1UxoEq1x4GXaG9ghBQJO_C85p3ZPU3T7zeKhy0Ij-UA4');
  Logger.log('Autorización concedida');
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

    // Busca la primera pregunta de escala (la del 1-5)
    var items = form.getItems();
    var npsIdx = -1;
    for (var i = 0; i < items.length; i++) {
      var t = items[i].getType();
      if (t === FormApp.ItemType.LINEAR_SCALE || t === FormApp.ItemType.SCALE) {
        npsIdx = i; break;
      }
    }

    // Escala 1-5: 5=Promotor, 4=Pasivo, 1-3=Detractor
    var promotores = 0, pasivos = 0, detractores = 0;
    if (npsIdx >= 0) {
      mesRes.forEach(function(r) {
        var ir = r.getItemResponses();
        if (ir[npsIdx]) {
          var score = parseInt(ir[npsIdx].getResponse(), 10);
          if (score === 5)      promotores++;
          else if (score === 4) pasivos++;
          else if (score >= 1)  detractores++;
        }
      });
    }

    var total = mesRes.length;
    return {
      ok: true,
      totalMes:    total,
      promotores:  promotores,
      pasivos:     pasivos,
      detractores: detractores,
      nps: total > 0 ? Math.round((promotores / total - detractores / total) * 100) : null
    };
  } catch(e) { return { ok: false, error: e.toString() }; }
}

