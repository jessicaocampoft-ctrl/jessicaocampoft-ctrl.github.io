// Sesión aislada para el editor del Pasaporte.
// Este archivo es aditivo: no cambia cloudflareSession ni el flujo del administrador.

function passportEditorConstantTimeEqual_(a, b) {
  a = String(a || '');
  b = String(b || '');
  var difference = a.length ^ b.length;
  var length = Math.max(a.length, b.length);
  for (var i = 0; i < length; i++) {
    difference |= (i < a.length ? a.charCodeAt(i) : 0) ^
      (i < b.length ? b.charCodeAt(i) : 0);
  }
  return difference === 0;
}

function passportEditorSignature_(message, secret) {
  var bytes = Utilities.computeHmacSha256Signature(
    message,
    secret,
    Utilities.Charset.UTF_8
  );
  return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/, '');
}

function createPassportEditorSession_(p) {
  var email = String(p.email || '').trim().toLowerCase();
  var role = String(p.role || '').trim().toUpperCase();
  var timestamp = String(p.timestamp || '');
  var nonce = String(p.nonce || '');
  var signature = String(p.signature || '');
  var secret = PropertiesService.getScriptProperties()
    .getProperty('PASSPORT_EDITOR_BRIDGE_SECRET') || '';
  var timestampNumber = Number(timestamp);

  if (
    !secret ||
    !email ||
    !/^[0-9]{10,16}$/.test(timestamp) ||
    !/^[A-Za-z0-9_-]{20,200}$/.test(nonce)
  ) {
    return {ok:false, error:'Solicitud de identidad no válida'};
  }

  if (Math.abs(Date.now() - timestampNumber) > 60000) {
    return {ok:false, error:'Solicitud de identidad vencida'};
  }

  if (['ADMIN', 'AUXILIAR'].indexOf(role) === -1) {
    return {ok:false, error:'Rol no autorizado'};
  }

  var message = email + '\n' + role + '\n' + timestamp + '\n' + nonce;
  var expected = passportEditorSignature_(message, secret);
  if (!passportEditorConstantTimeEqual_(expected, signature)) {
    return {ok:false, error:'Firma de identidad no válida'};
  }

  var cache = CacheService.getScriptCache();
  var nonceKey = 'passport_editor_nonce_' + nonce;
  if (cache.get(nonceKey)) {
    return {ok:false, error:'Solicitud de identidad repetida'};
  }
  cache.put(nonceKey, '1', 120);

  return {
    ok: true,
    sessionToken: createSession(),
    identity: {email: email, role: role}
  };
}

// Apps Script compila todos los archivos del proyecto en el mismo ámbito global.
// Conservamos el doGet actual y solo interceptamos la acción nueva.
var passportEditorOriginalDoGet_ = doGet;
doGet = function(e) {
  var p = e && e.parameter ? e.parameter : {};
  if (p.action === 'passportEditorSession') {
    return js(createPassportEditorSession_(p));
  }
  return passportEditorOriginalDoGet_(e);
};
