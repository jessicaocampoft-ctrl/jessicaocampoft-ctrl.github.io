// =============================================================
//  JESSICA OCAMPO FISIOTERAPEUTA — Apps Script Backend
//  Funciones: Reservas, Base de datos, Disponibilidad,
//             Panel Admin, Recordatorios diarios
// =============================================================

var ADMIN_TOKEN  = 'JESSICA2026';          // Cambia esta contrasena
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
    return js(getAvailability(p.date));
  }

  if (p.token !== ADMIN_TOKEN) {
    return js({ok: false, error: 'Sin permiso'});
  }

  if (p.action === 'adminData')    return js(getAdminData());
  if (p.action === 'block')        return js(doBlock(p));
  if (p.action === 'unblock')      return js(doUnblock(p));
  if (p.action === 'updateStatus') return js(doUpdateStatus(p));
  if (p.action === 'adminBook')    return js(createBooking(JSON.parse(decodeURIComponent(p.data)), true));

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
function createBooking(d, isAdmin) {
  if (!isAdmin) {
    var avail = checkAvailability(d.date, d.time, d.modality);
    if (!avail.available) return {ok: false, error: avail.reason};
  }

  var cal   = CalendarApp.getDefaultCalendar();
  var start = parseDT(d.date, d.time);
  // Domicilio para pacientes = 60min sesion + 30min buffer transporte
  var mins  = (d.modality === 'Domicilio' && !isAdmin) ? 90 : 60;
  var end   = new Date(start.getTime() + mins * 60000);
  var price = d.modality === 'Presencial' ? d.priceP : d.priceD;

  var event = cal.createEvent('[CITA] ' + d.service + ' - ' + d.name, start, end, {
    description: buildDesc(d, price),
    location: d.modality === 'Domicilio' ? (d.address || 'Domicilio - Pereira / Dosquebradas') : 'Pereira, Colombia'
  });
  event.addEmailReminder(60);
  event.addPopupReminder(30);

  // Guardar en Google Sheets
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

  // Link de WhatsApp para Jessica
  var tel  = (d.phone || '').replace(/\D/g,'');
  if (tel.length <= 10) tel = '57' + tel;
  var waConfirm = 'Hola ' + d.name + ', te confirmo tu cita de ' + d.service + ' el ' + d.date + ' a las ' + d.time + '. Quedo pendiente! - Jessica Ocampo Fisioterapeuta';
  var waLink = 'https://wa.me/' + tel + '?text=' + encodeURIComponent(waConfirm);

  // Email a Jessica con link de WhatsApp
  GmailApp.sendEmail(
    JESSICA_EMAIL,
    'Nueva cita: ' + d.name + ' - ' + d.service + ' | ' + d.date,
    buildEmailJessica(d, price) + '\n\n>> Confirmar al paciente por WhatsApp (1 clic):\n' + waLink + '\n\nID cita: ' + id
  );

  // Confirmacion al cliente
  if (d.email && d.email.indexOf('@') > 0) {
    GmailApp.sendEmail(d.email,
      'Solicitud de cita recibida - Jessica Ocampo Fisioterapeuta',
      buildEmailCliente(d, price)
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
function getAvailability(date) {
  var SLOTS = ['07:00','08:00','09:00','10:00','11:00','13:00','14:00','15:00','16:00','17:00','18:00','19:00'];
  var result = {};

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
    var start  = parseDT(date, s);
    var end60  = new Date(start.getTime() + 60 * 60000);
    var ok     = true;

    // 1. Google Calendar (eventos personales + citas existentes)
    for (var k = 0; k < calEvents.length && ok; k++) {
      var ev = calEvents[k];
      if (ev.isAllDayEvent()) continue;
      if (start < ev.getEndTime() && end60 > ev.getStartTime()) ok = false;
    }

    // 2. Citas en Sheets
    for (var i = 1; i < cRows.length && ok; i++) {
      var r = cRows[i];
      if (r[10] === 'Cancelada') continue;
      var rf = sd(r[7]);
      if (rf !== date) continue;
      var es = parseDT(rf, st(r[8]));
      var em = r[6] === 'Domicilio' ? 90 : 60;
      if (start < new Date(es.getTime() + em*60000) && end60 > es) ok = false;
    }

    // 3. Bloqueos
    for (var j = 1; j < bRows.length && ok; j++) {
      var b = bRows[j];
      if (sd(b[0]) !== date) continue;
      if (start < parseDT(date, st(b[2])) && end60 > parseDT(date, st(b[1]))) ok = false;
    }

    result[s] = ok;
  });

  return {ok: true, date: date, slots: result};
}

function checkAvailability(date, time, modality) {
  var start = parseDT(date, time);
  var mins  = modality === 'Domicilio' ? 90 : 60;
  var end   = new Date(start.getTime() + mins * 60000);

  // 1. Google Calendar — bloquear si hay cualquier evento personal
  try {
    var calEvents = CalendarApp.getDefaultCalendar().getEvents(start, end);
    for (var k = 0; k < calEvents.length; k++) {
      if (!calEvents[k].isAllDayEvent()) return {available: false, reason: 'Ese horario no esta disponible. Por favor elige otro.'};
    }
  } catch(x) {}

  var ss = getOrCreateSheet();

  // 2. Citas existentes en Sheets
  var cRows = ss.getSheetByName('Citas').getDataRange().getValues();
  for (var i = 1; i < cRows.length; i++) {
    var r = cRows[i];
    if (r[10] === 'Cancelada') continue;
    var rf = sd(r[7]);
    if (rf !== date) continue;
    var es = parseDT(rf, st(r[8]));
    var em = r[6] === 'Domicilio' ? 90 : 60;
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
  getOrCreateSheet().getSheetByName('Bloqueos').appendRow([
    p.date, p.startTime, p.endTime, p.reason || 'Bloqueado', 'Admin'
  ]);
  return {ok: true};
}

function doUnblock(p) {
  var sheet = getOrCreateSheet().getSheetByName('Bloqueos');
  var rows  = sheet.getDataRange().getValues();
  for (var i = rows.length - 1; i >= 1; i--) {
    if (rows[i][0] === p.date && rows[i][1] === p.startTime) {
      sheet.deleteRow(i + 1);
      return {ok: true};
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

function getAdminData() {
  var ss = getOrCreateSheet();

  var cRows = ss.getSheetByName('Citas').getDataRange().getValues();
  var citas = [];
  for (var i = 1; i < cRows.length; i++) {
    var r = cRows[i];
    citas.push({
      id: r[0], fechaReg: r[1],
      nombre: r[2],
      telefono: (r[3] instanceof Error || r[3] === null || r[3] === undefined) ? '' : ('' + r[3]),
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
      fecha: (b[0] instanceof Date) ? fmtDate(b[0]) : (b[0] ? ('' + b[0]).split('T')[0] : ''),
      inicio: (b[1] instanceof Date) ? (pad(b[1].getHours()) + ':' + pad(b[1].getMinutes())) : ('' + (b[1] || '')),
      fin: (b[2] instanceof Date) ? (pad(b[2].getHours()) + ':' + pad(b[2].getMinutes())) : ('' + (b[2] || '')),
      motivo: b[3]
    });
  }

  return {ok: true, citas: citas, bloqueos: bloqueos};
}

// -------------------------------------------------------------
//  RECORDATORIOS DIARIOS — ejecutar con trigger 7am
// -------------------------------------------------------------
function sendReminders() {
  var ss   = getOrCreateSheet();
  var rows = ss.getSheetByName('Citas').getDataRange().getValues();

  var today    = fmtDate(new Date());
  var tmrwDate = new Date(); tmrwDate.setDate(tmrwDate.getDate() + 1);
  var tomorrow = fmtDate(tmrwDate);

  var linksHoy = [], linksMañana = [];

  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (r[10] === 'Cancelada') continue;
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
      // Email recordatorio manana
      if (email && email.indexOf('@') > 0) {
        GmailApp.sendEmail(email,
          'Recordatorio: tu cita de manana - Jessica Ocampo Fisioterapeuta',
          'Hola ' + nombre + ',\n\nTe recordamos que manana tienes cita:\n\n' +
          'Servicio: ' + serv + '\nFecha: ' + fecha + '\nHora: ' + hora + '\nModalidad: ' + mod + '\nValor: ' + precio +
          '\n\nPara cancelar o reprogramar escribe al +57 313 646 7945.\n\n- Jessica Ocampo Fisioterapeuta\njessicaocampoft-ctrl.github.io'
        );
      }
      var msg1 = 'Hola ' + nombre + '! Te recuerdo que MANANA tienes cita de ' + serv + ' a las ' + hora + '. Cualquier cambio avisame! - Jessica';
      linksMañana.push(nombre + ' (' + hora + '): https://wa.me/' + phone + '?text=' + encodeURIComponent(msg1));
    }

    if (fecha === today) {
      // Email recordatorio hoy
      if (email && email.indexOf('@') > 0) {
        GmailApp.sendEmail(email,
          'Recordatorio: tienes cita HOY - Jessica Ocampo Fisioterapeuta',
          'Hola ' + nombre + ',\n\nTe recordamos que HOY tienes cita:\n\n' +
          'Servicio: ' + serv + '\nHora: ' + hora + '\nModalidad: ' + mod +
          '\n\nCualquier consulta: +57 313 646 7945\n\n- Jessica Ocampo Fisioterapeuta'
        );
      }
      var msg2 = 'Hola ' + nombre + '! Hoy tienes tu cita de ' + serv + ' a las ' + hora + '. Nos vemos! - Jessica';
      linksHoy.push(nombre + ' (' + hora + '): https://wa.me/' + phone + '?text=' + encodeURIComponent(msg2));
    }
  }

  if (linksHoy.length > 0 || linksMañana.length > 0) {
    var body = 'Recordatorios automaticos del dia ' + today + '\n\n';
    if (linksHoy.length)   body += '== CITAS DE HOY (WhatsApp 1 clic) ==\n' + linksHoy.join('\n') + '\n\n';
    if (linksMañana.length) body += '== CITAS DE MANANA (WhatsApp 1 clic) ==\n' + linksMañana.join('\n') + '\n';
    GmailApp.sendEmail(JESSICA_EMAIL, 'Recordatorios de citas - ' + today, body);
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
  if (files.hasNext()) return SpreadsheetApp.open(files.next());

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
  return ss;
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
  return 'Hola ' + d.name + ',\n\nTu solicitud de cita fue recibida.\n\n' +
    'SERVICIO: '   + d.service  + '\nFecha: '     + d.date +
    '\nHora: '     + d.time     + '\nModalidad: ' + d.modality + '\nValor: ' + price +
    '\n\nJessica te confirmara tu cita pronto por WhatsApp al +57 313 646 7945.\n' +
    'Si necesitas hacer algun cambio, escribele directamente.\n\n' +
    '- Jessica Ocampo\nFisioterapeuta — Pereira, Colombia\njessicaocampoft-ctrl.github.io';
}
