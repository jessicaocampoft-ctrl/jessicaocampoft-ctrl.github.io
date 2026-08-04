function base64Url(bytes) {
  let value = '';
  for (const byte of bytes) value += String.fromCharCode(byte);
  return btoa(value).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}

function decodeBase64Url(value) {
  const normalized = value.replace(/-/g, '+').replace(/_/g, '/');
  const binary = atob(normalized + '='.repeat((4 - normalized.length % 4) % 4));
  return Uint8Array.from(binary, (char) => char.charCodeAt(0));
}

function roleFor(email, env) {
  const normalized = String(email || '').toLowerCase();
  if (normalized === String(env.ADMIN_EMAIL || '').toLowerCase()) return 'ADMIN';
  if (normalized === String(env.AUXILIAR_EMAIL || '').toLowerCase()) return 'AUXILIAR';
  return null;
}

async function accessEmail(request, env) {
  const token = request.headers.get('Cf-Access-Jwt-Assertion');
  if (!token || !env.ACCESS_TEAM_DOMAIN || !env.ACCESS_AUD) return null;
  const parts = token.split('.');
  if (parts.length !== 3) return null;
  try {
    const header = JSON.parse(new TextDecoder().decode(decodeBase64Url(parts[0])));
    const claims = JSON.parse(new TextDecoder().decode(decodeBase64Url(parts[1])));
    const certs = await (await fetch(`https://${env.ACCESS_TEAM_DOMAIN}/cdn-cgi/access/certs`)).json();
    const key = certs.keys.find((item) => item.kid === header.kid);
    const audience = Array.isArray(claims.aud) ? claims.aud : [claims.aud];
    if (!key || !audience.includes(env.ACCESS_AUD) || Number(claims.exp) * 1000 <= Date.now()) return null;
    const publicKey = await crypto.subtle.importKey('jwk', key, { name: 'RSASSA-PKCS1-v1_5', hash: 'SHA-256' }, false, ['verify']);
    const signed = new TextEncoder().encode(`${parts[0]}.${parts[1]}`);
    if (!await crypto.subtle.verify('RSASSA-PKCS1-v1_5', publicKey, decodeBase64Url(parts[2]), signed)) return null;
    return typeof claims.email === 'string' ? claims.email.toLowerCase() : null;
  } catch (_) { return null; }
}

async function requireEditor(request, env) {
  const email = await accessEmail(request, env);
  const role = roleFor(email, env);
  return role ? { email, role } : null;
}

async function editorSession(auth, env) {
  if (!env.PASSPORT_EDITOR_BRIDGE_SECRET || !env.APPS_SCRIPT_URL) throw new Error('Configuración segura incompleta.');
  const timestamp = String(Date.now());
  const nonce = crypto.randomUUID().replace(/-/g, '') + base64Url(crypto.getRandomValues(new Uint8Array(16)));
  const canonical = `${auth.email}\n${auth.role}\n${timestamp}\n${nonce}`;
  const key = await crypto.subtle.importKey('raw', new TextEncoder().encode(env.PASSPORT_EDITOR_BRIDGE_SECRET), { name: 'HMAC', hash: 'SHA-256' }, false, ['sign']);
  const signature = base64Url(new Uint8Array(await crypto.subtle.sign('HMAC', key, new TextEncoder().encode(canonical))));
  const url = new URL(env.APPS_SCRIPT_URL);
  url.searchParams.set('action', 'passportEditorSession');
  url.searchParams.set('email', auth.email);
  url.searchParams.set('role', auth.role);
  url.searchParams.set('timestamp', timestamp);
  url.searchParams.set('nonce', nonce);
  url.searchParams.set('signature', signature);
  const response = await fetch(url, { redirect: 'follow', cache: 'no-store' });
  const payload = await response.json();
  if (!response.ok || !payload.ok || !payload.sessionToken) throw new Error(payload.error || 'No fue posible iniciar la sesión del editor.');
  return payload.sessionToken;
}

async function passportCall(auth, env, action, fields) {
  const token = await editorSession(auth, env);
  const url = new URL(env.APPS_SCRIPT_URL);
  url.searchParams.set('action', action);
  url.searchParams.set('token', token);
  for (const [key, value] of Object.entries(fields || {})) url.searchParams.set(key, typeof value === 'string' ? value : JSON.stringify(value));
  const response = await fetch(url, { redirect: 'follow', cache: 'no-store' });
  const payload = await response.json();
  if (!response.ok || !payload.ok) throw new Error(payload.error || 'No fue posible completar la operación.');
  return payload;
}

function json(value, status = 200) {
  return Response.json(value, { status, headers: { 'Cache-Control': 'no-store' } });
}

function stamps(values, maximum) {
  const source = values || {};
  const answer = {};
  for (let index = 1; index <= maximum; index += 1) answer[String(index)] = Boolean(source[String(index)] ?? source[index]);
  return answer;
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    if (!url.pathname.startsWith('/api/')) return env.ASSETS.fetch(request);
    const auth = await requireEditor(request, env);
    if (!auth) return json({ ok: false, error: 'No autorizado.' }, 403);
    try {
      if (url.pathname === '/api/session' && request.method === 'GET') return json({ ok: true, role: auth.role });
      if (request.method !== 'POST') return json({ ok: false, error: 'Método no permitido.' }, 405);
      const body = await request.json();
      if (url.pathname === '/api/passport') {
        const nombre = String(body.nombre || '').trim();
        const telefono = String(body.telefono || '').replace(/\D/g, '');
        if (!nombre || nombre.length > 120 || telefono.length > 20) return json({ ok: false, error: 'Nombre o teléfono no válido.' }, 400);
        return json(await passportCall(auth, env, 'passportEnsure', { nombre, telefono }));
      }
      if (url.pathname === '/api/save') {
        const id = String(body.id || '');
        if (!id || id.length > 100) return json({ ok: false, error: 'Pasaporte no válido.' }, 400);
        return json(await passportCall(auth, env, 'passportSaveProgress', {
          id,
          passport: { stamps: stamps(body.stamps, 16) },
          descarga: { stamps: stamps(body.reto, 2) }
        }));
      }
      if (url.pathname === '/api/regenerate') {
        const id = String(body.id || '');
        if (!id || id.length > 100) return json({ ok: false, error: 'Pasaporte no válido.' }, 400);
        return json(await passportCall(auth, env, 'passportRegenerateToken', { id }));
      }
      return json({ ok: false, error: 'Ruta no encontrada.' }, 404);
    } catch (error) {
      return json({ ok: false, error: error instanceof Error ? error.message : 'Error inesperado.' }, 502);
    }
  }
};

