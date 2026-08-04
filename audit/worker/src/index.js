function roleFor(email, env) {
  const value = String(email || '').toLowerCase();
  if (value === String(env.ADMIN_EMAIL || '').toLowerCase()) return 'ADMIN';
  if (value === String(env.AUXILIAR_EMAIL || '').toLowerCase()) return 'AUXILIAR';
  if (value === String(env.FISIOTERAPEUTA_EMAIL || '').toLowerCase()) return 'FISIOTERAPEUTA';
  return null;
}

function sessionResponse(email, env) {
  const rol = roleFor(email, env);
  if (!rol) return Response.json({ autorizado: false }, { status: 403, headers: { 'Cache-Control': 'no-store' } });
  return Response.json({ email, rol, autorizado: true }, { headers: { 'Cache-Control': 'no-store' } });
}

function decodeBase64Url(value) {
  const normalized = value.replace(/-/g, '+').replace(/_/g, '/');
  const padded = normalized + '='.repeat((4 - normalized.length % 4) % 4);
  const binary = atob(padded);
  return Uint8Array.from(binary, (character) => character.charCodeAt(0));
}

async function accessEmail(request, env) {
  const token = request.headers.get('Cf-Access-Jwt-Assertion');
  const teamDomain = env.ACCESS_TEAM_DOMAIN || env.TEAM_DOMAIN;
  const audienceId = env.ACCESS_AUD || env.POLICY_AUD;
  if (!token || !teamDomain || !audienceId) return null;
  const parts = token.split('.');
  if (parts.length !== 3) return null;
  try {
    const header = JSON.parse(new TextDecoder().decode(decodeBase64Url(parts[0])));
    const claims = JSON.parse(new TextDecoder().decode(decodeBase64Url(parts[1])));
    const certificates = await (await fetch(`https://${teamDomain}/cdn-cgi/access/certs`)).json();
    const key = certificates.keys.find((item) => item.kid === header.kid);
    const audience = Array.isArray(claims.aud) ? claims.aud : [claims.aud];
    if (!key || claims.exp * 1000 <= Date.now() || !audience.includes(audienceId)) return null;
    const publicKey = await crypto.subtle.importKey('jwk', key, { name: 'RSASSA-PKCS1-v1_5', hash: 'SHA-256' }, false, ['verify']);
    const payload = new TextEncoder().encode(`${parts[0]}.${parts[1]}`);
    if (!await crypto.subtle.verify('RSASSA-PKCS1-v1_5', publicKey, decodeBase64Url(parts[2]), payload)) return null;
    return typeof claims.email === 'string' ? claims.email : null;
  } catch (_) {
    return null;
  }
}

function base64Url(bytes) {
  let binary = '';
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}

function corsHeaders(request) {
  const origin = request.headers.get('Origin');
  return origin === 'https://admin.cuidandotefisioterapia.com'
    ? { 'Access-Control-Allow-Origin': origin, 'Access-Control-Allow-Credentials': 'true', Vary: 'Origin' }
    : {};
}

async function bootstrap(request, env, email) {
  const role = roleFor(email, env);
  if (role !== 'ADMIN' && role !== 'AUXILIAR') {
    return Response.json({ autorizado: false }, { status: 403, headers: corsHeaders(request) });
  }
  if (!env.BRIDGE_SECRET || !env.APPS_SCRIPT_URL) {
    return Response.json({ autorizado: false }, { status: 503, headers: corsHeaders(request) });
  }

  const timestamp = String(Date.now());
  const nonce = crypto.randomUUID().replace(/-/g, '') + base64Url(crypto.getRandomValues(new Uint8Array(16)));
  const canonical = `${email.toLowerCase()}\n${role}\n${timestamp}\n${nonce}`;
  const key = await crypto.subtle.importKey(
    'raw',
    new TextEncoder().encode(env.BRIDGE_SECRET),
    { name: 'HMAC', hash: 'SHA-256' },
    false,
    ['sign']
  );
  const signature = base64Url(new Uint8Array(await crypto.subtle.sign('HMAC', key, new TextEncoder().encode(canonical))));
  const upstream = new URL(env.APPS_SCRIPT_URL);
  upstream.searchParams.set('action', 'cloudflareSession');
  upstream.searchParams.set('email', email.toLowerCase());
  upstream.searchParams.set('role', role);
  upstream.searchParams.set('timestamp', timestamp);
  upstream.searchParams.set('nonce', nonce);
  upstream.searchParams.set('signature', signature);

  try {
    const response = await fetch(upstream, { redirect: 'follow', cache: 'no-store' });
    const data = await response.json();
    return Response.json(data, {
      status: response.ok ? 200 : 502,
      headers: { 'Cache-Control': 'no-store', ...corsHeaders(request) }
    });
  } catch (_) {
    return Response.json({ autorizado: false }, { status: 502, headers: corsHeaders(request) });
  }
}

async function requireAdmin_(request, env) {
  const email = await accessEmail(request, env);
  const role = roleFor(email, env);
  return role === 'ADMIN' || role === 'AUXILIAR' ? { email, role } : null;
}

async function appsProxy_(request, env) {
  const source = new URL(request.url);
  const target = new URL(env.APPS_SCRIPT_URL);
  target.search = source.search;
  const init = { method: request.method, redirect: 'follow', cache: 'no-store' };
  if (request.method !== 'GET' && request.method !== 'HEAD') {
    init.body = await request.text();
    init.headers = { 'Content-Type': request.headers.get('Content-Type') || 'application/json' };
  }
  const upstream = await fetch(target, init);
  return new Response(upstream.body, {
    status: upstream.status,
    headers: {
      'Content-Type': upstream.headers.get('Content-Type') || 'application/json',
      'Cache-Control': 'no-store'
    }
  });
}

async function testAdmin_(request, env, email) {
  const bootstrapResponse = await bootstrap(request, env, email);
  if (!bootstrapResponse.ok) {
    return new Response('No fue posible iniciar la sesión de pruebas.', {
      status: 502,
      headers: { 'Cache-Control': 'no-store' }
    });
  }

  const bootstrapPayload = await bootstrapResponse.json();
  if (!bootstrapPayload || !bootstrapPayload.ok || !bootstrapPayload.sessionToken) {
    return new Response('No autorizado.', { status: 403, headers: { 'Cache-Control': 'no-store' } });
  }

  const assetResponse = await env.ASSETS.fetch(new Request(new URL('/admin-live.html', request.url)));
  if (!assetResponse.ok) {
    return new Response('Respaldo del administrador no disponible.', {
      status: 500,
      headers: { 'Cache-Control': 'no-store' }
    });
  }

  let html = await assetResponse.text();
  if (!/const APPS_SCRIPT_URL = '[^']+';/.test(html)) {
    return new Response('Respaldo incompatible: falta la configuración del backend.', { status: 500 });
  }
  if (!html.includes("if (TOKEN) {\n    const btn = document.getElementById('loginBtn');")) {
    return new Response('Respaldo incompatible: falta el inicio administrativo.', { status: 500 });
  }

  const safePayload = JSON.stringify(bootstrapPayload).replace(/</g, '\\u003c');
  const boot = '<script>(function(){var payload=' + safePayload + ';sessionStorage.setItem("adminToken",payload.sessionToken);window.__BOOTSTRAP_ADMIN_TEST__=true;window.addEventListener("DOMContentLoaded",async function(){var login=document.getElementById("loginScreen"),app=document.getElementById("adminApp"),error=document.getElementById("loginErr"),button=document.getElementById("loginBtn");try{TOKEN=payload.sessionToken;allData=payload;_loginTime=Date.now();await loadAdminKV();await loadTeamData();await limpiarHorariosInvalidosAuto();reloadMetas();_initSidebarState();initDashboard();await _runUrlRepairIfRequested();login.style.display="none";app.style.display="block";}catch(e){if(app)app.style.display="none";if(login)login.style.display="block";if(error){error.textContent=(e&&e.message)||"No fue posible cargar el panel de pruebas.";error.style.display="block";}if(button){button.disabled=false;button.textContent="Reintentar";button.onclick=function(){location.reload();};}}});})();</script>';

  html = html.replace(/LogoCuidandote\//g, 'https://admin.cuidandotefisioterapia.com/LogoCuidandote/');
  html = html.replace('admin-copy-tools.js', 'assets/admin-copy-tools.js');
  html = html.replace(/const APPS_SCRIPT_URL = '[^']+';/, "const APPS_SCRIPT_URL = location.origin + '/api';");
  html = html.replace(
    "if (TOKEN) {\n    const btn = document.getElementById('loginBtn');",
    "if (TOKEN && !window.__BOOTSTRAP_ADMIN_TEST__) {\n    const btn = document.getElementById('loginBtn');"
  );
  html = html.replace(
    '</head>',
    '<style>#userInput,#pwInput,label[for="userInput"],label[for="pwInput"]{display:none!important}</style></head>'
  );

  const bridge = '<script src="/bridge.js" defer></script>';
  html = html.replace(/<\/body>\s*<\/html>\s*$/i, boot + bridge + '</body></html>');
  return new Response(html, {
    headers: { 'Content-Type': 'text/html; charset=UTF-8', 'Cache-Control': 'no-store' }
  });
}

function bridgeScript_() {
  return `
(function () {
  function showPassportError(message) {
    if (typeof toast === 'function') toast(message, 'err');
    else window.alert(message);
  }

  function openPassportEditor() {
    if (typeof renderPasaporteAdminTools !== 'function') {
      showPassportError('El editor del Pasaporte no está disponible.');
      return;
    }

    renderPasaporteAdminTools();
    var tools = document.getElementById('pasAdminTools');
    if (!tools) {
      showPassportError('No fue posible abrir el editor. Vuelve a seleccionar el paciente.');
      return;
    }

    tools.style.display = 'block';
    tools.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  function patchPassportButton() {
    var button = document.getElementById('pasAbrirBtn');
    if (!button) return;
    button.textContent = 'Editar progreso →';
    button.type = 'button';
    button.onclick = function (event) {
      if (event) event.preventDefault();
      openPassportEditor();
      return false;
    };
  }

  function startBridge() {
    window.logout = function () {
      sessionStorage.removeItem('adminToken');
      sessionStorage.removeItem('adminUser');
      location.assign('/logout');
    };

    window.abrirPasaporte = openPassportEditor;
    patchPassportButton();

    var observer = new MutationObserver(function () {
      patchPassportButton();
    });
    observer.observe(document.body, { childList: true, subtree: true });
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', startBridge);
  } else {
    startBridge();
  }
})();`;
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    if (url.pathname === '/logout') {
      return new Response(null, {
        status: 302,
        headers: { Location: '/cdn-cgi/access/logout', 'Cache-Control': 'no-store' }
      });
    }

    if (url.pathname === '/admin-test') {
      const auth = await requireAdmin_(request, env);
      if (!auth) return Response.json({ autorizado: false }, { status: 403 });
      return testAdmin_(request, env, auth.email);
    }

    if (url.pathname === '/bridge.js') {
      if (!await requireAdmin_(request, env)) return new Response('', { status: 403 });
      return new Response(bridgeScript_(), {
        headers: {
          'Content-Type': 'application/javascript; charset=UTF-8',
          'Cache-Control': 'no-store'
        }
      });
    }

    if (url.pathname === '/api') {
      if (!await requireAdmin_(request, env)) {
        return Response.json({ autorizado: false }, { status: 403 });
      }
      return appsProxy_(request, env);
    }

    const email = await accessEmail(request, env);
    if (!email) {
      return Response.json({ autorizado: false }, {
        status: 401,
        headers: { 'Cache-Control': 'no-store' }
      });
    }

    if (url.pathname === '/bootstrap' && request.method === 'GET') {
      return bootstrap(request, env, email);
    }
    if (url.pathname === '/session' && request.method === 'GET') {
      return sessionResponse(email, env);
    }

    const rol = roleFor(email, env);
    if (url.pathname === '/admin') {
      if (rol !== 'ADMIN' && rol !== 'AUXILIAR') {
        return new Response('No autorizado', { status: 403, headers: { 'Cache-Control': 'no-store' } });
      }
      return new Response('Acceso administrativo autorizado', {
        headers: { 'Cache-Control': 'no-store' }
      });
    }

    if (url.pathname === '/fisioterapeuta') {
      if (rol !== 'FISIOTERAPEUTA') {
        return new Response('No autorizado', { status: 403, headers: { 'Cache-Control': 'no-store' } });
      }
      return new Response('Acceso de fisioterapeuta autorizado', {
        headers: { 'Cache-Control': 'no-store' }
      });
    }

    return new Response('No encontrado', { status: 404 });
  }
};