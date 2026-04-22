// =============================================================
//  JESSICA OCAMPO FISIOTERAPEUTA — Apps Script Backend
//  Funciones: Reservas, Base de datos, Disponibilidad,
//             Panel Admin, Recordatorios diarios
// =============================================================

var ADMIN_TOKEN  = 'JESSICA2026';          // Cambia esta contrasena
var JESSICA_EMAIL = 'jessica.ocampo.ft@gmail.com';
var JESSICA_WA    = '573136467945';
var SS_NAME       = 'Citas Jessica Ocampo Fisio';
var GEMINI_API_KEY = 'AIzaSyBZ0o7H6eIrDewQ_SSYSNMKlqMzY75xvjM';

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

  if (p.token !== ADMIN_TOKEN) {
    return js({ok: false, error: 'Sin permiso'});
  }

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

  // Pasaporte — escritura (requiere token admin)
  if (p.action === 'savePassport' && p.nombre) {
    return js(savePassport(decodeURIComponent(p.nombre), p.passport || '{}', p.descarga || '{}'));
  }

  return txt('Jessica Ocampo Fisioterapeuta - Sistema activo');
}

// -------------------------------------------------------------
//  POST — Reservas de pacientes
// -------------------------------------------------------------
function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);
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
  var price = d.modality === 'Presencial' ? d.priceP : d.priceD;

  // Solo crear evento en Google Calendar si es una cita real (no un registro)
  if (!soloRegistro) {
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
  }

  // Para registros de paciente: solo guardar en hoja Pacientes, sin cita ni email
  if (soloRegistro) {
    upsertPaciente(d.name, d.phone, d.email);
    return {ok: true, id: 'REG-' + new Date().getTime()};
  }

  // Guardar en Google Sheets (solo citas reales)
  var id    = 'C' + new Date().getTime();
  var ss    = getOrCreateSheet();
  var cSheet = ss.getSheetByName('Citas');
  var phoneClean = ('' + (d.phone||'')).replace(/\D/g,'');
  cSheet.appendRow([
    id,
    new Date().toLocaleString('es-CO'),
    d.name, phoneClean, d.email,
    d.service, d.modality,
    d.date, d.time, price,
    'Confirmada',
    d.address || '', d.notes || '', ''
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
  var waConfirm = 'Hola ' + d.name + ', te confirmo tu cita de ' + d.service + ' el ' + d.date + ' a las ' + d.time + '. Quedo pendiente! - Jessica Ocampo Fisioterapeuta';
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
    if (d.notas !== undefined) sheet.getRange(i+1, 13).setValue(d.notas);
    try {
      var dp = oldFecha.split('-');
      var dayS = new Date(+dp[0], +dp[1]-1, +dp[2], 0, 0, 0);
      var dayE = new Date(+dp[0], +dp[1]-1, +dp[2], 23, 59, 59);
      var calEvs = CalendarApp.getDefaultCalendar().getEvents(dayS, dayE);
      for (var k = 0; k < calEvs.length; k++) {
        var t = calEvs[k].getTitle() || '';
        if (t.indexOf('[CITA]') === 0 && t.indexOf(rows[i][2]) > -1) {
          var ns = parseDT(d.fecha || oldFecha, d.hora || oldHora);
          calEvs[k].setTime(ns, new Date(ns.getTime() + 60*60000));
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
    if (cRows[i][2] === nombre) cSheet.deleteRow(i + 1);
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
    if (cRows[i][2] !== d.oldNombre) continue;
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
      hora: (r[8] instanceof Date) ? (pad(r[8].getHours()) + ':' + pad(r[8].getMinutes())) : ('' + (r[8] || '')),
      precio: r[9],
      estado: r[10], direccion: r[11], notas: r[12], notaAdmin: r[13]
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

  return {ok: true, citas: citas, bloqueos: bloqueos, pacientes: pacientes};
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
    var hora  = (r[8] instanceof Date) ? (pad(r[8].getHours()) + ':' + pad(r[8].getMinutes())) :
                (typeof r[8] === 'number') ? (pad(Math.floor(r[8]*24)) + ':' + pad(Math.round((r[8]*24%1)*60))) :
                ('' + (r[8]||''));
    var rawTel = (r[3] instanceof Error) ? '' : ('' + (r[3]||''));
    var phone  = rawTel.replace(/\D/g,'');
    if (phone.length <= 10) phone = '57' + phone;

    if (fecha === tomorrow) {
      if (email && email.indexOf('@') > 0) {
        var dp = fecha.split('-');
        var fObj = new Date(+dp[0], +dp[1]-1, +dp[2]);
        var fechaLegible = diasSemana[fObj.getDay()] + ' ' + +dp[2] + ' de ' + meses[+dp[1]-1];
        GmailApp.sendEmail(
          email,
          'Recordatorio: mañana tienes cita — Jessica Ocampo Fisioterapeuta',
          'Este correo requiere un cliente de correo con soporte HTML.',
          {htmlBody: buildReminderEmail(nombre, serv, fechaLegible, hora, mod, precio, false),
           name: 'Jessica Ocampo Fisioterapeuta'}
        );
      }
      var msg1 = 'Hola ' + nombre + '! Te recuerdo que mañana tienes cita de ' + serv + ' a las ' + hora + '. Cualquier cambio avísame! - Jessica';
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
          {htmlBody: buildReminderEmail(nombre, serv, fechaLegible2, hora, mod, precio, true),
           name: 'Jessica Ocampo Fisioterapeuta'}
        );
      }
      var msg2 = 'Hola ' + nombre + '! Hoy tienes tu cita de ' + serv + ' a las ' + hora + '. Nos vemos! - Jessica';
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
  Logger.log('Trigger activado: sendReminders cada dia a las 7am hora Colombia.');
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
  cs.getRange(1,1,1,14).setValues([[
    'ID','FechaRegistro','Nombre','Telefono','Email',
    'Servicio','Modalidad','FechaCita','Hora','Precio',
    'Estado','Direccion','Notas','NotaAdmin'
  ]]);
  ss.insertSheet('Bloqueos').getRange(1,1,1,5).setValues([[
    'Fecha','HoraInicio','HoraFin','Motivo','CreadoPor'
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
  if (s.indexOf('completa') > -1)       return 80;
  if (s.indexOf('readaptacion') > -1)   return 50;
  if (s.indexOf('valoracion') > -1)     return 50;
  if (s.indexOf('piernas') > -1)        return 50;
  if (s.indexOf('cuello') > -1)         return 50;
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

  var html =
    '<div style="font-family:Arial,sans-serif;max-width:520px;margin:0 auto;border:1px solid #e5e7eb;border-radius:12px;overflow:hidden">' +
    '<div style="background:#0d9488;padding:28px 32px;text-align:center">' +
    '<h1 style="color:#fff;margin:0;font-size:20px">✅ Cita Confirmada</h1>' +
    '<p style="color:#ccfbf1;margin:6px 0 0;font-size:14px">Jessica Ocampo Fisioterapeuta</p>' +
    '</div>' +
    '<div style="padding:28px 32px">' +
    '<p style="margin:0 0 20px;font-size:15px">Hola <strong>' + d.name + '</strong>,<br>tu cita está <strong>confirmada</strong>. Aquí están los detalles:</p>' +
    '<table style="width:100%;border-collapse:collapse;font-size:14px">' +
    '<tr><td style="padding:10px 0;border-bottom:1px solid #f3f4f6;color:#6b7280;width:120px">Servicio</td><td style="padding:10px 0;border-bottom:1px solid #f3f4f6;font-weight:600">' + d.service + '</td></tr>' +
    '<tr><td style="padding:10px 0;border-bottom:1px solid #f3f4f6;color:#6b7280">Fecha</td><td style="padding:10px 0;border-bottom:1px solid #f3f4f6;font-weight:600">' + fechaLegible + '</td></tr>' +
    '<tr><td style="padding:10px 0;border-bottom:1px solid #f3f4f6;color:#6b7280">Hora</td><td style="padding:10px 0;border-bottom:1px solid #f3f4f6;font-weight:600">' + d.time + '</td></tr>' +
    '<tr><td style="padding:10px 0;border-bottom:1px solid #f3f4f6;color:#6b7280">Modalidad</td><td style="padding:10px 0;border-bottom:1px solid #f3f4f6">' + d.modality + (d.address ? ' — ' + d.address : '') + '</td></tr>' +
    '<tr><td style="padding:10px 0;color:#6b7280">Valor</td><td style="padding:10px 0;font-weight:600">' + price + '</td></tr>' +
    '</table>' +
    '<div style="background:#f0fdf4;border-radius:8px;padding:14px 18px;margin:20px 0;font-size:13px;color:#166534">' +
    '📅 Recibirás un recordatorio por correo el día anterior y el mismo día de tu cita.' +
    '</div>' +
    '<div style="background:#fffbeb;border:1px solid #fde68a;border-radius:8px;padding:16px 18px;margin:16px 0">' +
    '<p style="margin:0 0 10px;font-size:13px;font-weight:700;color:#92400e">💳 Formas de pago</p>' +
    '<table style="width:100%;font-size:13px;color:#44403c;border-collapse:collapse">' +
    '<tr><td style="padding:4px 0;color:#78716c;width:110px">Bancolombia</td><td style="padding:4px 0;font-weight:600">Cta. Ahorros · 91257857099</td></tr>' +
    '<tr><td style="padding:4px 0;color:#78716c">Llave</td><td style="padding:4px 0;font-weight:600">1010124692</td></tr>' +
    '<tr><td style="padding:4px 0;color:#78716c">Nequi</td><td style="padding:4px 0;font-weight:600">3136467945</td></tr>' +
    '<tr><td style="padding:4px 0;color:#78716c">A nombre de</td><td style="padding:4px 0">Jessica Andrea Ocampo Barbosa</td></tr>' +
    '</table>' +
    '</div>' +
    '<p style="font-size:13px;color:#6b7280;margin:0">¿Necesitas cancelar o cambiar? Escríbele directamente:<br>' +
    '<a href="https://wa.me/573136467945" style="color:#0d9488">+57 313 646 7945 (WhatsApp)</a></p>' +
    '</div>' +
    '<div style="background:#f9fafb;padding:16px 32px;text-align:center;font-size:12px;color:#9ca3af">' +
    'Jessica Ocampo Fisioterapeuta · Pereira, Colombia<br>' +
    '<a href="https://jessicaocampoft-ctrl.github.io" style="color:#0d9488">jessicaocampoft-ctrl.github.io</a>' +
    '</div></div>';

  return html;
}

function buildReminderEmail(nombre, serv, fechaLegible, hora, mod, precio, esHoy) {
  var titulo = esHoy ? '⏰ Hoy tienes cita' : '📅 Recordatorio: mañana tienes cita';
  var intro  = esHoy
    ? '¡Hola <strong>' + nombre + '</strong>! Hoy es el día de tu cita. Aquí te recordamos los detalles:'
    : 'Hola <strong>' + nombre + '</strong>, mañana tienes tu cita. Te recordamos los detalles:';

  return '<div style="font-family:Arial,sans-serif;max-width:520px;margin:0 auto;border:1px solid #e5e7eb;border-radius:12px;overflow:hidden">' +
    '<div style="background:' + (esHoy ? '#0284c7' : '#0d9488') + ';padding:24px 32px;text-align:center">' +
    '<h1 style="color:#fff;margin:0;font-size:19px">' + titulo + '</h1>' +
    '<p style="color:#e0f2fe;margin:6px 0 0;font-size:13px">Jessica Ocampo Fisioterapeuta</p>' +
    '</div>' +
    '<div style="padding:28px 32px">' +
    '<p style="margin:0 0 20px;font-size:15px">' + intro + '</p>' +
    '<table style="width:100%;border-collapse:collapse;font-size:14px">' +
    '<tr><td style="padding:10px 0;border-bottom:1px solid #f3f4f6;color:#6b7280;width:120px">Servicio</td><td style="padding:10px 0;border-bottom:1px solid #f3f4f6;font-weight:600">' + serv + '</td></tr>' +
    '<tr><td style="padding:10px 0;border-bottom:1px solid #f3f4f6;color:#6b7280">Fecha</td><td style="padding:10px 0;border-bottom:1px solid #f3f4f6;font-weight:600">' + fechaLegible + '</td></tr>' +
    '<tr><td style="padding:10px 0;border-bottom:1px solid #f3f4f6;color:#6b7280">Hora</td><td style="padding:10px 0;border-bottom:1px solid #f3f4f6;font-weight:600">' + hora + '</td></tr>' +
    '<tr><td style="padding:10px 0;color:#6b7280">Modalidad</td><td style="padding:10px 0">' + mod + '</td></tr>' +
    (precio ? '<tr><td style="padding:10px 0;color:#6b7280">Valor</td><td style="padding:10px 0">' + precio + '</td></tr>' : '') +
    '</table>' +
    '<div style="background:#fffbeb;border:1px solid #fde68a;border-radius:8px;padding:14px 18px;margin:16px 0">' +
    '<p style="margin:0 0 8px;font-size:13px;font-weight:700;color:#92400e">💳 Formas de pago</p>' +
    '<table style="width:100%;font-size:13px;color:#44403c;border-collapse:collapse">' +
    '<tr><td style="padding:3px 0;color:#78716c;width:110px">Bancolombia</td><td style="padding:3px 0;font-weight:600">Cta. Ahorros · 91257857099</td></tr>' +
    '<tr><td style="padding:3px 0;color:#78716c">Llave</td><td style="padding:3px 0;font-weight:600">1010124692</td></tr>' +
    '<tr><td style="padding:3px 0;color:#78716c">Nequi</td><td style="padding:3px 0;font-weight:600">3136467945</td></tr>' +
    '<tr><td style="padding:3px 0;color:#78716c">A nombre de</td><td style="padding:3px 0">Jessica Andrea Ocampo Barbosa</td></tr>' +
    '</table>' +
    '</div>' +
    '<p style="font-size:13px;color:#6b7280;margin:0">¿Necesitas cancelar o cambiar? Escríbele directamente:<br>' +
    '<a href="https://wa.me/573136467945" style="color:#0d9488">+57 313 646 7945 (WhatsApp)</a></p>' +
    '</div>' +
    '<div style="background:#f9fafb;padding:16px 32px;text-align:center;font-size:12px;color:#9ca3af">' +
    'Jessica Ocampo Fisioterapeuta · Pereira, Colombia' +
    '</div></div>';
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
    '<p style="font-size:13px;color:#6b7280;margin:0">O escríbeme directamente:<br>' +
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

function generateEvalReport(d) {
  try {
    if (!GEMINI_API_KEY || GEMINI_API_KEY === 'PEGA_AQUI_TU_CLAVE_GEMINI') {
      return { ok: false, error: 'Configura GEMINI_API_KEY en el script. Obtenla gratis en aistudio.google.com' };
    }

    var prompt = buildEvalPrompt(d);
    var url    = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=' + GEMINI_API_KEY;

    var response = UrlFetchApp.fetch(url, {
      method:      'post',
      contentType: 'application/json',
      muteHttpExceptions: true,
      payload: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { temperature: 0.7, maxOutputTokens: 1200 }
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
    if (sc.balanceMono) {
      screensText += '• Balance Monopodal: D=' + (sc.balanceMono.timeD || 'No eval') + 's' +
        ', I=' + (sc.balanceMono.timeI || 'No eval') + 's';
      if (sc.balanceMono.obs) screensText += ' — ' + sc.balanceMono.obs;
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
    '\nGenera el reporte en español con EXACTAMENTE estas secciones (usa los títulos en mayúsculas y negrilla):\n\n' +
    '**RESUMEN EJECUTIVO**\n' +
    '[2-3 oraciones sobre el estado postural y funcional general. Tono profesional y empático. Menciona el impacto directo en el rendimiento CrossFit.]\n\n' +
    '**ANÁLISIS POR ZONAS**\n' +
    '[Para cada zona con hallazgos, explica en 1-2 oraciones qué significa biomecánicamente y cómo afecta los movimientos específicos de CrossFit (snatch, clean, squat, deadlift, etc.). Si no hay hallazgos en una zona, omítela.]\n\n' +
    '**ANÁLISIS FUNCIONAL**\n' +
    '[Interpreta los resultados de las pruebas funcionales: qué patrones de movimiento están comprometidos, qué músculos o cadenas están débiles o rígidas, y cómo se correlacionan con los hallazgos posturales. Sé específica con los scores bajos.]\n\n' +
    '**RIESGOS IDENTIFICADOS**\n' +
    '[Lista máximo 4 riesgos concretos de lesión si no se atienden. Específicos para CrossFit. Usa viñetas con •]\n\n' +
    '**PLAN DE ACCIÓN RECOMENDADO**\n' +
    '[3-5 recomendaciones concretas ordenadas por prioridad. Incluye qué tipo de trabajo se haría en fisioterapia. Usa viñetas con •]\n\n' +
    '**CONCLUSIÓN**\n' +
    '[1 párrafo motivador. Menciona que con un plan de fisioterapia deportiva personalizado se pueden corregir estos disbalances y mejorar el rendimiento. Invita a dar el siguiente paso. No menciones precios.]\n\n' +
    'REGLAS: Máximo 500 palabras en total. Sé directa y específica. Usa lenguaje claro, no excesivamente técnico. Tono: profesional pero cercano.';
}

