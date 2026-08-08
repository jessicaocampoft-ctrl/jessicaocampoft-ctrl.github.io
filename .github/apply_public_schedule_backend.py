from pathlib import Path
import re

backend_path = Path('google-apps-script.js')
html_path = Path('index.html')
backend = backend_path.read_text(encoding='utf-8-sig')
html = html_path.read_text(encoding='utf-8-sig')

# 1) Acción pública para leer configuración.
reviews_block = """  if (p.action === 'getReviews') {\n    return js(getGoogleReviews());\n  }\n"""
public_action = reviews_block + """\n  // Horarios públicos por sede — lectura pública, sin datos sensibles.\n  if (p.action === 'publicScheduleConfig') {\n    return js(getPublicScheduleConfig_());\n  }\n"""
if "p.action === 'publicScheduleConfig'" not in backend:
    if reviews_block not in backend:
        raise SystemExit('No se encontró el bloque getReviews para insertar publicScheduleConfig')
    backend = backend.replace(reviews_block, public_action, 1)

# 2) Acción autenticada POST para guardar desde administrador.
post_marker = """    if (d.action === 'savePayment') {\n      if (!validateSession(d.token)) return js({ok: false, error: 'Sin permiso'});\n"""
post_insert = """    if (d.action === 'savePublicScheduleConfig') {\n      if (!validateSession(d.token)) return js({ok: false, error: 'Sin permiso'});\n      var scheduleUser = getSessionUser_(d.token) || {id:'admin', nombre:'Administracion', rol:'Superadministradora'};\n      return js(savePublicScheduleConfig_(d.data || {}, scheduleUser));\n    }\n"""
if "d.action === 'savePublicScheduleConfig'" not in backend:
    if post_marker not in backend:
        raise SystemExit('No se encontró savePayment en doPost para insertar savePublicScheduleConfig')
    backend = backend.replace(post_marker, post_insert + post_marker, 1)

# 3) Helpers de configuración. Se guardan en ScriptProperties para compartirlos entre despliegues.
helpers = r'''
// =============================================================
//  HORARIOS PÚBLICOS POR SEDE — configuración compartida Admin/Web
// =============================================================
var PUBLIC_SCHEDULE_CONFIG_KEY_ = 'PUBLIC_SCHEDULE_CONFIG_V1';

function defaultPublicWeekly_() {
  return {
    '0': [],
    '1': [['08:00','16:30']],
    '2': [['08:00','17:00']],
    '3': [['08:00','17:00']],
    '4': [['08:00','20:00']],
    '5': [['08:00','20:00']],
    '6': [['07:00','09:30']]
  };
}

function defaultPublicScheduleConfig_() {
  return {
    version: 1,
    venues: {
      santa: {
        label: 'Sede Santa Mónica',
        enabled: true,
        services: [
          'Descarga Muscular — Cuello y Espalda',
          'Descarga Muscular — Piernas',
          'Descarga Muscular Completa',
          'Valoración Funcional',
          'Readaptación Funcional'
        ],
        weekly: defaultPublicWeekly_()
      },
      recovery: {
        label: 'Sede Campestre Recovery',
        enabled: true,
        services: [
          'Descarga Muscular — Cuello y Espalda',
          'Descarga Muscular — Piernas',
          'Descarga Muscular Completa'
        ],
        weekly: defaultPublicWeekly_()
      }
    },
    updatedAt: ''
  };
}

function validPublicTime_(value) {
  return /^(?:[01]\d|2[0-3]):[0-5]\d$/.test('' + (value || ''));
}

function normalizePublicRanges_(ranges) {
  if (!Array.isArray(ranges)) return [];
  var clean = [];
  for (var i = 0; i < ranges.length && clean.length < 2; i++) {
    var pair = ranges[i];
    if (!Array.isArray(pair) || pair.length < 2) continue;
    var start = '' + pair[0];
    var end = '' + pair[1];
    if (!validPublicTime_(start) || !validPublicTime_(end)) continue;
    if (minutesFromTime_(start) >= minutesFromTime_(end)) continue;
    clean.push([start, end]);
  }
  clean.sort(function(a,b){ return minutesFromTime_(a[0]) - minutesFromTime_(b[0]); });
  var result = [];
  for (var j = 0; j < clean.length; j++) {
    if (result.length && minutesFromTime_(clean[j][0]) < minutesFromTime_(result[result.length-1][1])) continue;
    result.push(clean[j]);
  }
  return result;
}

function sanitizePublicScheduleConfig_(input) {
  var defaults = defaultPublicScheduleConfig_();
  var source = input;
  if (typeof source === 'string') {
    try { source = JSON.parse(source); } catch(e) { source = {}; }
  }
  if (!source || typeof source !== 'object') source = {};
  var sourceVenues = source.venues || {};
  var allowed = {
    santa: defaults.venues.santa.services.slice(),
    recovery: defaults.venues.recovery.services.slice()
  };
  var out = {version:1, venues:{}, updatedAt: source.updatedAt || ''};

  ['santa','recovery'].forEach(function(key) {
    var def = defaults.venues[key];
    var src = sourceVenues[key] || {};
    var srcServices = Array.isArray(src.services) ? src.services : def.services;
    var services = srcServices.filter(function(service) { return allowed[key].indexOf(service) >= 0; });
    var weekly = {};
    for (var day = 0; day <= 6; day++) {
      var dayKey = '' + day;
      if (src.weekly && Object.prototype.hasOwnProperty.call(src.weekly, dayKey)) {
        weekly[dayKey] = normalizePublicRanges_(src.weekly[dayKey]);
      } else {
        weekly[dayKey] = normalizePublicRanges_(def.weekly[dayKey]);
      }
    }
    out.venues[key] = {
      label: def.label,
      enabled: typeof src.enabled === 'boolean' ? src.enabled : def.enabled,
      services: services,
      weekly: weekly
    };
  });
  return out;
}

function getPublicScheduleConfig_() {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty(PUBLIC_SCHEDULE_CONFIG_KEY_);
    if (!raw) return {ok:true, config:defaultPublicScheduleConfig_(), source:'default'};
    var parsed = JSON.parse(raw);
    return {ok:true, config:sanitizePublicScheduleConfig_(parsed), source:'saved'};
  } catch(e) {
    return {ok:true, config:defaultPublicScheduleConfig_(), source:'fallback'};
  }
}

function savePublicScheduleConfig_(input, user) {
  try {
    var config = sanitizePublicScheduleConfig_(input);
    config.updatedAt = new Date().toISOString();
    PropertiesService.getScriptProperties().setProperty(PUBLIC_SCHEDULE_CONFIG_KEY_, JSON.stringify(config));
    try { auditTeam_(user, 'Actualizó horarios públicos', '', '', '', JSON.stringify({venues:config.venues,updatedAt:config.updatedAt})); } catch(auditError) {}
    return {ok:true, config:config};
  } catch(e) {
    return {ok:false, error:'No se pudieron guardar los horarios públicos: ' + e.message};
  }
}

function publicVenueKey_(modality) {
  var value = ('' + (modality || '')).toLowerCase();
  if (value.indexOf('santa') >= 0) return 'santa';
  if (value.indexOf('recovery') >= 0 || value.indexOf('campestre') >= 0) return 'recovery';
  return 'domicilio';
}

function configuredPublicRanges_(date, modality, service) {
  var key = publicVenueKey_(modality);
  // Domicilios conserva la jornada pública histórica; este módulo controla las dos sedes físicas.
  if (key === 'domicilio') return null;
  var response = getPublicScheduleConfig_();
  var config = response && response.config;
  var venue = config && config.venues && config.venues[key];
  if (!venue || venue.enabled === false) return [];
  if (service && (!Array.isArray(venue.services) || venue.services.indexOf('' + service) < 0)) return [];
  var dp = ('' + date).split('-');
  var d = new Date(+dp[0], +dp[1]-1, +dp[2], 0, 0, 0);
  return normalizePublicRanges_(venue.weekly && venue.weekly['' + d.getDay()] || []);
}

function publicCandidateSlotsConfigured_(date, durationMins, modality, service) {
  var ranges = configuredPublicRanges_(date, modality, service);
  if (ranges === null) return publicCandidateSlots_(date, durationMins);
  var out = [];
  ranges.forEach(function(range) {
    var start = minutesFromTime_(range[0]);
    var close = minutesFromTime_(range[1]);
    for (var mins = start; mins + durationMins <= close; mins += 60) out.push(timeFromMinutes_(mins));
  });
  return out;
}

function fitsPublicScheduleConfigured_(date, time, durationMins, modality, service) {
  var ranges = configuredPublicRanges_(date, modality, service);
  if (ranges === null) return fitsPublicSchedule_(date, time, durationMins);
  var start = minutesFromTime_(time);
  var end = start + durationMins;
  return ranges.some(function(range) {
    return start >= minutesFromTime_(range[0]) && end <= minutesFromTime_(range[1]);
  });
}
'''

if 'PUBLIC_SCHEDULE_CONFIG_KEY_' not in backend:
    marker = "function getAvailability(date, service, modality) {"
    if marker not in backend:
        raise SystemExit('No se encontró getAvailability')
    backend = backend.replace(marker, helpers + "\n" + marker, 1)

# 4) Availability devuelve solamente slots habilitados para la sede/servicio.
old_avail = """function getAvailability(date, service, modality) {\n  var SLOTS = publicCandidateSlots_(date, getServiceDuration(service) + (modality === 'Domicilio' ? 30 : 0));\n  var result = {};\n  var newDur = getServiceDuration(service) + (modality === 'Domicilio' ? 30 : 0); // duraci"""
if old_avail in backend:
    # Conservar el comentario original sin depender de su codificación posterior.
    backend = backend.replace(
        "function getAvailability(date, service, modality) {\n  var SLOTS = publicCandidateSlots_(date, getServiceDuration(service) + (modality === 'Domicilio' ? 30 : 0));\n  var result = {};\n  var newDur = getServiceDuration(service) + (modality === 'Domicilio' ? 30 : 0);",
        "function getAvailability(date, service, modality) {\n  var newDur = getServiceDuration(service) + (modality === 'Domicilio' ? 30 : 0);\n  var SLOTS = publicCandidateSlotsConfigured_(date, newDur, modality, service);\n  var result = {};",
        1
    )
elif 'publicCandidateSlotsConfigured_(date, newDur, modality, service)' not in backend:
    raise SystemExit('No se pudo actualizar getAvailability')

# 5) La validación final de una reserva pública usa la configuración por sede.
old_fit = "if (!fitsPublicSchedule_(date, time, mins)) {"
if old_fit in backend:
    # La última definición de validatePublicBookingSchedule_ prevalece; sustituir todas las comprobaciones públicas es seguro.
    backend = backend.replace(old_fit, "if (!fitsPublicScheduleConfigured_(date, time, mins, modality, service)) {")

# 6) Cargar el cliente de horarios al final de la web, después del script de reservas existente.
script_tag = '  <script src="public-schedule-client.js?v=20260807-1"></script>\n'
if 'public-schedule-client.js' not in html:
    if '</body>' not in html:
        raise SystemExit('index.html no tiene </body>')
    html = html.replace('</body>', script_tag + '</body>', 1)

# Validaciones de seguridad.
required_backend = [
    "p.action === 'publicScheduleConfig'",
    "d.action === 'savePublicScheduleConfig'",
    'PUBLIC_SCHEDULE_CONFIG_KEY_',
    'publicCandidateSlotsConfigured_',
    'fitsPublicScheduleConfigured_',
    'PropertiesService.getScriptProperties().setProperty(PUBLIC_SCHEDULE_CONFIG_KEY_'
]
for token in required_backend:
    if token not in backend:
        raise SystemExit('Falta backend: ' + token)
for token in ['getPassportSecure', 'passportSaveProgress_', 'getGoogleReviews', 'createBooking']:
    if token not in backend:
        raise SystemExit('Se perdió integración crítica backend: ' + token)
for token in ['public-schedule-client.js', 'APPS_SCRIPT_URL', 'loadReviews();', 'pasaporte.html']:
    if token not in html:
        raise SystemExit('Se perdió integración crítica web: ' + token)

backend_path.write_text('\ufeff' + backend, encoding='utf-8')
html_path.write_text(html, encoding='utf-8')
print('Backend y web pública de horarios preparados correctamente.')
