var __defProp = Object.defineProperty;
var __name = (target, value) => __defProp(target, "name", { value, configurable: true });

// src/index.js
function roleFor(email, env) {
  const value = String(email || "").toLowerCase();
  if (value === String(env.ADMIN_EMAIL || "").toLowerCase()) return "ADMIN";
  if (value === String(env.AUXILIAR_EMAIL || "").toLowerCase()) return "AUXILIAR";
  if (value === String(env.FISIOTERAPEUTA_EMAIL || "").toLowerCase()) return "FISIOTERAPEUTA";
  return null;
}
__name(roleFor, "roleFor");
function sessionResponse(email, env) {
  const rol = roleFor(email, env);
  if (!rol) return Response.json({ autorizado: false }, { status: 403, headers: { "Cache-Control": "no-store" } });
  return Response.json({ email, rol, autorizado: true }, { headers: { "Cache-Control": "no-store" } });
}
__name(sessionResponse, "sessionResponse");
function decodeBase64Url(value) {
  const normalized = value.replace(/-/g, "+").replace(/_/g, "/");
  const padded = normalized + "=".repeat((4 - normalized.length % 4) % 4);
  const binary = atob(padded);
  return Uint8Array.from(binary, (character) => character.charCodeAt(0));
}
__name(decodeBase64Url, "decodeBase64Url");
async function accessEmail(request, env) {
  const token = request.headers.get("Cf-Access-Jwt-Assertion");
  const teamDomain = env.ACCESS_TEAM_DOMAIN || env.TEAM_DOMAIN;
  const audienceId = env.ACCESS_AUD || env.POLICY_AUD;
  if (!token || !teamDomain || !audienceId) return null;
  const parts = token.split(".");
  if (parts.length !== 3) return null;
  try {
    const header = JSON.parse(new TextDecoder().decode(decodeBase64Url(parts[0])));
    const claims = JSON.parse(new TextDecoder().decode(decodeBase64Url(parts[1])));
    const certificates = await (await fetch(`https://${teamDomain}/cdn-cgi/access/certs`)).json();
    const key = certificates.keys.find((item) => item.kid === header.kid);
    const audience = Array.isArray(claims.aud) ? claims.aud : [claims.aud];
    if (!key || claims.exp * 1e3 <= Date.now() || !audience.includes(audienceId)) return null;
    const publicKey = await crypto.subtle.importKey("jwk", key, { name: "RSASSA-PKCS1-v1_5", hash: "SHA-256" }, false, ["verify"]);
    const payload = new TextEncoder().encode(`${parts[0]}.${parts[1]}`);
    if (!await crypto.subtle.verify("RSASSA-PKCS1-v1_5", publicKey, decodeBase64Url(parts[2]), payload)) return null;
    return typeof claims.email === "string" ? claims.email : null;
  } catch (_) {
    return null;
  }
}
__name(accessEmail, "accessEmail");
function base64Url(bytes) {
  let binary = "";
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}
__name(base64Url, "base64Url");
function corsHeaders(request) {
  const origin = request.headers.get("Origin");
  return origin === "https://admin.cuidandotefisioterapia.com" ? { "Access-Control-Allow-Origin": origin, "Access-Control-Allow-Credentials": "true", Vary: "Origin" } : {};
}
__name(corsHeaders, "corsHeaders");
async function bootstrap(request, env, email) {
  const role = roleFor(email, env);
  if (role !== "ADMIN" && role !== "AUXILIAR") return Response.json({ autorizado: false }, { status: 403, headers: corsHeaders(request) });
  if (!env.BRIDGE_SECRET || !env.APPS_SCRIPT_URL) return Response.json({ autorizado: false }, { status: 503, headers: corsHeaders(request) });
  const timestamp = String(Date.now());
  const nonce = crypto.randomUUID().replace(/-/g, "") + base64Url(crypto.getRandomValues(new Uint8Array(16)));
  const canonical = `${email.toLowerCase()}
${role}
${timestamp}
${nonce}`;
  const key = await crypto.subtle.importKey("raw", new TextEncoder().encode(env.BRIDGE_SECRET), { name: "HMAC", hash: "SHA-256" }, false, ["sign"]);
  const signature = base64Url(new Uint8Array(await crypto.subtle.sign("HMAC", key, new TextEncoder().encode(canonical))));
  const upstream = new URL(env.APPS_SCRIPT_URL);
  upstream.searchParams.set("action", "cloudflareSession");
  upstream.searchParams.set("email", email.toLowerCase());
  upstream.searchParams.set("role", role);
  upstream.searchParams.set("timestamp", timestamp);
  upstream.searchParams.set("nonce", nonce);
  upstream.searchParams.set("signature", signature);
  try {
    const response = await fetch(upstream, { redirect: "follow", cache: "no-store" });
    const data = await response.json();
    return Response.json(data, { status: response.ok ? 200 : 502, headers: { "Cache-Control": "no-store", ...corsHeaders(request) } });
  } catch (_) {
    return Response.json({ autorizado: false }, { status: 502, headers: corsHeaders(request) });
  }
}
__name(bootstrap, "bootstrap");
async function requireAdmin_(request, env) {
  const email = await accessEmail(request, env);
  const role = roleFor(email, env);
  return role === "ADMIN" || role === "AUXILIAR" ? { email, role } : null;
}
__name(requireAdmin_, "requireAdmin_");
async function appsProxy_(request, env) {
  const source = new URL(request.url);
  const target = new URL(env.APPS_SCRIPT_URL);
  target.search = source.search;
  const init = { method: request.method, redirect: "follow", cache: "no-store" };
  if (request.method !== "GET" && request.method !== "HEAD") {
    init.body = await request.text();
    init.headers = { "Content-Type": request.headers.get("Content-Type") || "application/json" };
  }
  const upstream = await fetch(target, init);
  return new Response(upstream.body, { status: upstream.status, headers: { "Content-Type": upstream.headers.get("Content-Type") || "application/json", "Cache-Control": "no-store" } });
}
__name(appsProxy_, "appsProxy_");
async function testAdmin_(request, env) {
  const upstream = await fetch("https://admin.cuidandotefisioterapia.com/", { cache: "no-store" });
  let html = await upstream.text();
  html = html.replace(/LogoCuidandote\//g, "https://admin.cuidandotefisioterapia.com/LogoCuidandote/");
  html = html.replace(/const APPS_SCRIPT_URL = '[^']+';/, "const APPS_SCRIPT_URL = location.origin + '/api';");
  html = html.replace(
    /function abrirPasaporte\(\) \{[\s\S]*?window\.open\(link, '_blank'\);\s*\}/,
    "function abrirPasaporte() { renderPasaporteAdminTools(); var tools = document.getElementById('pasAdminTools'); if (tools) tools.scrollIntoView({behavior:'smooth', block:'start'}); }"
  );
  html = html.replace("Abrir y editar \u2192", "Editar progreso \u2192");
  html = html.replace(
    /function abrirPasaporte\(\) \{[\s\S]*?window\.open\(link, '_blank'\);\s*\}/,
    "function abrirPasaporte() { renderPasaporteAdminTools(); var tools = document.getElementById('pasAdminTools'); if (tools) tools.scrollIntoView({behavior:'smooth', block:'start'}); }"
  );
  html = html.replace("Abrir y editar \u2192", "Editar progreso \u2192");
  html = html.replace("</head>", '<style>#userInput,#pwInput,label[for="userInput"],label[for="pwInput"]{display:none!important}</style></head>');
  const bridge = '<script src="/bridge.js" defer><\/script>';
  html = html.replace(/<\/body>\s*<\/html>\s*$/i, bridge + "</body></html>");
  return new Response(html, { headers: { "Content-Type": "text/html; charset=UTF-8", "Cache-Control": "no-store" } });
}
__name(testAdmin_, "testAdmin_");
function bridgeScript_() {
  return `
  (function () {
  async function startBridge() {
    var button = document.getElementById('loginBtn');
    var error = document.getElementById('loginErr');
    if (button) { button.textContent = 'Ingresando con correo autorizado...'; button.disabled = true; }
    window.doLogin = async function () {
      if (button) { button.disabled = true; button.textContent = 'Verificando identidad...'; }
      if (error) error.style.display = 'none';
      try {
        var controller = new AbortController();
        var timeout = setTimeout(function () { controller.abort(); }, 15000);
        var response;
        try {
          response = await fetch('/bootstrap', { credentials: 'include', cache: 'no-store', signal: controller.signal });
        } finally { clearTimeout(timeout); }
        var payload;
        try { payload = await response.json(); }
        catch (jsonError) { throw new Error('El servidor devolvi\xF3 una respuesta no v\xE1lida.'); }
        if (!response.ok) throw new Error((payload && payload.error) || 'No fue posible validar la identidad.');
        if (!payload || !payload.ok || !payload.sessionToken) throw new Error((payload && payload.error) || 'No autorizado.');

        TOKEN = payload.sessionToken;
        sessionStorage.setItem('adminToken', TOKEN);
        sessionStorage.setItem('adminUser', JSON.stringify(payload.identity || {}));
        _loginTime = Date.now(); allData = payload;
        document.getElementById('loginScreen').style.display = 'none';
        document.getElementById('adminApp').style.display = 'block';
        reloadMetas(); _initSidebarState(); initDashboard(); await _runUrlRepairIfRequested();
        Promise.allSettled([loadAdminKV(), loadTeamData(), limpiarHorariosInvalidosAuto()]);
      } catch (loginError) {
        var message = loginError && loginError.name === 'AbortError'
          ? 'La verificaci\xF3n de identidad tard\xF3 demasiado. Intenta nuevamente.'
          : (loginError && /aborted/i.test(loginError.message || '')
            ? 'La verificaci\xF3n de identidad tard\xF3 demasiado. Intenta nuevamente.'
            : (loginError && loginError.message ? loginError.message : 'No fue posible validar tu identidad.'));
        if (error) { error.textContent = message; error.style.display = 'block'; }
      } finally { if (button) { button.disabled = false; button.textContent = 'Ingresar'; } }
    };
    window.logout = function () { sessionStorage.removeItem('adminToken'); sessionStorage.removeItem('adminUser'); location.assign('/logout'); };
    // El enlace p\xFAblico es solo para el paciente. El editor permanece dentro
    // del panel, usando la sesi\xF3n administrativa existente.
    window.abrirPasaporte = function () {
      if (typeof renderPasaporteAdminTools === 'function') renderPasaporteAdminTools();
      var tools = document.getElementById('pasAdminTools');
      if (tools) tools.scrollIntoView({ behavior: 'smooth', block: 'start' });
    };
    function labelPassportEditor() {
      var editButton = document.getElementById('pasAbrirBtn');
      if (editButton) editButton.textContent = 'Editar progreso \u2192';
    }
    labelPassportEditor(); setTimeout(labelPassportEditor, 1500);
    window.doLogin();
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', startBridge);
  else startBridge();
  })();`;
}
__name(bridgeScript_, "bridgeScript_");
var index_default = {
  async fetch(request, env) {
    const url = new URL(request.url);
    if (url.pathname === "/logout") {
      return new Response(null, { status: 302, headers: { Location: "/cdn-cgi/access/logout", "Cache-Control": "no-store" } });
    }
    if (url.pathname === "/admin-test") {
      if (!await requireAdmin_(request, env)) return Response.json({ autorizado: false }, { status: 403 });
      return testAdmin_(request, env);
    }
    if (url.pathname === "/bridge.js") {
      if (!await requireAdmin_(request, env)) return new Response("", { status: 403 });
      return new Response(bridgeScript_(), { headers: { "Content-Type": "application/javascript; charset=UTF-8", "Cache-Control": "no-store" } });
    }
    if (url.pathname === "/api") {
      if (!await requireAdmin_(request, env)) return Response.json({ autorizado: false }, { status: 403 });
      return appsProxy_(request, env);
    }
    const email = await accessEmail(request, env);
    if (!email) return Response.json({ autorizado: false }, { status: 401, headers: { "Cache-Control": "no-store" } });
    if (url.pathname === "/bootstrap" && request.method === "GET") return bootstrap(request, env, email);
    if (url.pathname === "/session" && request.method === "GET") return sessionResponse(email, env);
    const rol = roleFor(email, env);
    if (url.pathname === "/admin") {
      if (rol !== "ADMIN" && rol !== "AUXILIAR") return new Response("No autorizado", { status: 403, headers: { "Cache-Control": "no-store" } });
      return new Response("Acceso administrativo autorizado", { headers: { "Cache-Control": "no-store" } });
    }
    if (url.pathname === "/fisioterapeuta") {
      if (rol !== "FISIOTERAPEUTA") return new Response("No autorizado", { status: 403, headers: { "Cache-Control": "no-store" } });
      return new Response("Acceso de fisioterapeuta autorizado", { headers: { "Cache-Control": "no-store" } });
    }
    return new Response("No encontrado", { status: 404 });
  }
};
export {
  index_default as default
};
//# sourceMappingURL=index.js.map
