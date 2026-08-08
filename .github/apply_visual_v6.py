from pathlib import Path
import re

p = Path('index.html')
t = p.read_text(encoding='utf-8')

# Menú superior: limpio, sin Pasaporte ni Contacto.
nav_pattern = re.compile(r'<ul class="nav-links" id="navLinks">.*?</ul>', re.S)
nav_new = '''<ul class="nav-links" id="navLinks">
        <li><a href="#servicios">Fisioterapia y rehabilitación</a></li>
        <li><a href="#experiencias">Masajes y bienestar</a></li>
        <li><a href="#equipo">Equipo</a></li>
        <li><a href="#agenda">Agenda</a></li>
        <li><a href="#agenda" class="nav-cta">Agendar cita</a></li>
      </ul>'''
t, n = nav_pattern.subn(nav_new, t, count=1)
if n != 1:
    raise SystemExit('No se pudo actualizar el menú')

# Hero: conservar contenido actual, con el acento visual aprobado.
t = t.replace(
    '<div class="hero-badge hero-enabled-badge">\n            <span class="hero-enabled-check" aria-hidden="true">✓</span>\n            Servicio de fisioterapia habilitado\n          </div>',
    '<div class="hero-badge hero-enabled-badge">Fisioterapia en Pereira y Dosquebradas</div>',
    1
)
t = t.replace(
    '<h1>Recupera tu movimiento con atención fisioterapéutica personalizada.</h1>',
    '<h1>Recupera tu movimiento con <span class="hero-accent">atención fisioterapéutica personalizada.</span></h1>',
    1
)
t = t.replace(
    'Valoración fisioterapéutica, rehabilitación funcional y ejercicio terapéutico con atención en nuestra sede de Dosquebradas y a domicilio en Pereira y Dosquebradas.',
    'Realizamos valoración fisioterapéutica, rehabilitación funcional y ejercicio terapéutico, con atención en nuestra sede de Dosquebradas y a domicilio en Pereira y Dosquebradas.',
    1
)

# Habilitación: lenguaje comprensible para clientes.
t = t.replace(
    '<div><strong>Servicio habilitado en REPS</strong><span>Fisioterapia · Profesional independiente</span></div>',
    '<div><strong>Servicio de fisioterapia habilitado</strong><span>Atención habilitada en sede y a domicilio · Registro vigente en REPS</span></div>',
    1
)
t = t.replace(
    'Cuidándote Fisioterapia presta el servicio de fisioterapia como profesional independiente habilitada en el Registro Especial de Prestadores de Servicios de Salud (REPS).',
    'Servicio registrado oficialmente en REPS como profesional independiente.',
    1
)

css = r'''
<style id="visual-refresh-v6-2026-08-07">
  :root {
    --primary:#7bd9cf;
    --primary-hover:#62cec3;
    --primary-dark:#29968e;
    --text:#252c2b;
    --text-muted:#63706d;
    --border:rgba(92,201,190,.28);
    --font-heading:'Cormorant Garamond', Georgia, serif;
    --font-body:'Trebuchet MS','DM Sans',Arial,sans-serif;
    --font-mono:'Trebuchet MS','DM Sans',Arial,sans-serif;
  }

  body { background:#fff!important; font-family:var(--font-body)!important; color:var(--text)!important; }
  h1,h2,h3,h4,h5,h6,.hero h1,.section-title,.about-content h2,.services-header h2,.experiences-header h2,
  .review-name,.plan-name,.pasaporte-title { font-family:var(--font-heading)!important; font-style:normal!important; color:var(--text)!important; }

  /* Header: logo centrado y menú debajo */
  .nav {
    position:relative!important;
    top:auto!important;left:auto!important;right:auto!important;
    padding:22px 0 0!important;
    background:#fff!important;
    border-bottom:1px solid rgba(92,201,190,.18)!important;
    box-shadow:none!important;
    backdrop-filter:none!important;
    -webkit-backdrop-filter:none!important;
  }
  .nav.scrolled { padding:22px 0 0!important; background:#fff!important; }
  .nav-inner {
    display:flex!important;
    flex-direction:column!important;
    align-items:center!important;
    justify-content:center!important;
    gap:20px!important;
    max-width:1200px!important;
  }
  .brand-logo,.nav.scrolled .brand-logo {
    width:280px!important;
    max-width:72vw!important;
    display:flex!important;
    flex-direction:column!important;
    align-items:center!important;
  }
  .brand-logo img { width:100%!important; height:auto!important; filter:none!important; opacity:1!important; }
  .brand-logo .logo-byline {
    display:block!important;
    font-family:var(--font-heading)!important;
    font-size:1rem!important;
    font-style:italic!important;
    color:#52b8ae!important;
    margin-top:2px!important;
    letter-spacing:.01em!important;
  }
  .nav-links {
    width:100%!important;
    min-height:72px!important;
    display:flex!important;
    justify-content:center!important;
    align-items:center!important;
    gap:clamp(30px,4.5vw,62px)!important;
    border-top:1px solid rgba(92,201,190,.12)!important;
  }
  .nav-links a {
    font-family:var(--font-heading)!important;
    font-size:1.17rem!important;
    font-weight:600!important;
    letter-spacing:.01em!important;
    text-transform:none!important;
    color:#343938!important;
    white-space:nowrap!important;
  }
  .nav-links a::after { display:none!important; }
  .nav-links a:hover,.nav-links a.active { color:var(--primary-dark)!important; }
  .nav-cta {
    background:var(--primary)!important;
    color:#29413e!important;
    padding:12px 24px!important;
    border-radius:999px!important;
    border:0!important;
    box-shadow:0 9px 22px rgba(123,217,207,.20)!important;
    font-family:var(--font-heading)!important;
    font-size:1.13rem!important;
  }

  /* Hero: degradado azul/turquesa suave, texto protagonista y foto proporcionada */
  .hero {
    min-height:auto!important;
    display:block!important;
    padding:72px 0 82px!important;
    background:
      radial-gradient(circle at 12% 20%,rgba(210,235,242,.58),transparent 34%),
      radial-gradient(circle at 88% 18%,rgba(198,239,234,.48),transparent 31%),
      linear-gradient(105deg,#fff 0%,#f7fbfd 47%,#edf8f7 100%)!important;
  }
  .hero::before { display:none!important; }
  .hero-grid {
    grid-template-columns:minmax(0,1.42fr) minmax(340px,.88fr)!important;
    gap:78px!important;
    align-items:center!important;
  }
  .hero-content { padding:12px 0!important; }
  .hero-enabled-badge {
    display:inline-flex!important;
    width:max-content!important;
    padding:8px 15px!important;
    margin-bottom:28px!important;
    font-family:var(--font-body)!important;
    font-size:.86rem!important;
    text-transform:none!important;
    letter-spacing:0!important;
    font-weight:700!important;
    color:var(--primary-dark)!important;
    background:rgba(255,255,255,.76)!important;
    border:1px solid rgba(92,201,190,.32)!important;
  }
  .hero-enabled-badge::before {
    content:'';
    width:9px;height:9px;border-radius:50%;background:var(--primary);
  }
  .hero h1 {
    max-width:760px!important;
    margin:0 0 26px!important;
    font-family:var(--font-heading)!important;
    font-size:clamp(4rem,5.6vw,5.9rem)!important;
    line-height:.98!important;
    font-weight:600!important;
    letter-spacing:-.035em!important;
    color:#242827!important;
  }
  .hero-accent {
    display:block!important;
    color:var(--primary)!important;
    font-family:inherit!important;
    font-style:normal!important;
    font-weight:500!important;
  }
  .hero-subtitle {
    max-width:690px!important;
    margin:0 0 30px!important;
    font-size:1.03rem!important;
    line-height:1.82!important;
    color:var(--text-muted)!important;
  }
  .hero-buttons { gap:14px!important; margin-bottom:0!important; }
  .btn-solid {
    background:var(--primary)!important;
    color:#243c39!important;
    border:0!important;
    border-radius:14px!important;
    min-height:54px!important;
    padding:0 25px!important;
    box-shadow:none!important;
  }
  .btn-outline {
    background:#fff!important;
    color:#247d75!important;
    border:1px solid rgba(92,201,190,.42)!important;
    border-radius:14px!important;
    min-height:54px!important;
    padding:0 25px!important;
  }
  .hero-trust-note {
    margin:23px 0 0!important;
    color:#697572!important;
    font-size:.78rem!important;
    line-height:1.5!important;
  }
  .hero-urgency,.hero-stats,.hero-photo-chip { display:none!important; }
  .hero-photo-wrap { display:flex!important; justify-content:center!important; padding:32px 0!important; }
  .hero-photo-frame {
    width:min(100%,440px)!important;
    max-width:440px!important;
    height:445px!important;
    padding:12px!important;
    border:1px solid rgba(92,201,190,.32)!important;
    border-radius:38px!important;
    background:rgba(255,255,255,.52)!important;
    box-shadow:0 22px 58px rgba(50,105,100,.10)!important;
    overflow:hidden!important;
  }
  .hero-photo-frame img {
    width:100%!important;
    height:100%!important;
    min-height:0!important;
    object-fit:cover!important;
    object-position:center 30%!important;
    border-radius:28px!important;
  }
  .hero-photo-frame .corner,.hero-photo-frame .overlay { display:none!important; }

  /* Habilitación: claro para cualquier cliente */
  .enablement-strip { margin-top:0!important; padding:0 0 80px!important; background:#f7fcfb!important; }
  .enablement-grid { border-radius:18px!important; box-shadow:0 14px 36px rgba(30,110,102,.06)!important; }
  .enablement-item { padding:22px 26px!important; }
  .enablement-item strong { color:#2b6e69!important; font-size:.88rem!important; }
  .enablement-item span { color:#71807d!important; font-size:.78rem!important; line-height:1.45!important; }
  .enablement-legal { color:#7d8986!important; font-size:.72rem!important; text-align:center!important; }

  /* Mantener V2 del resto, pero recuperar la tipografía de marca */
  .services-header h2,.experiences-header h2,.about-content h2,.reviews-header h2,
  .service-card h3,.experience-card h3,.plan-name,.pasaporte-title {
    font-family:var(--font-heading)!important;
    font-style:normal!important;
    font-weight:600!important;
  }

  @media(max-width:900px) {
    .hero-grid { grid-template-columns:1fr!important; gap:42px!important; }
    .hero-content { text-align:center!important; }
    .hero-enabled-badge { margin-left:auto!important; margin-right:auto!important; }
    .hero h1,.hero-subtitle { margin-left:auto!important; margin-right:auto!important; }
    .hero-buttons { justify-content:center!important; }
    .hero-trust-note { text-align:center!important; }
    .hero-photo-wrap { padding:0!important; }
  }

  @media(max-width:768px) {
    .nav { padding:12px 0!important; position:relative!important; }
    .nav-inner {
      flex-direction:row!important;
      justify-content:space-between!important;
      gap:14px!important;
    }
    .brand-logo,.nav.scrolled .brand-logo { width:165px!important; align-items:flex-start!important; }
    .brand-logo .logo-byline { font-size:.76rem!important; }
    .nav-links {
      position:fixed!important;
      top:0!important;right:-100%!important;
      width:78%!important;max-width:330px!important;height:100vh!important;
      min-height:0!important;
      flex-direction:column!important;
      justify-content:center!important;
      gap:28px!important;
      padding:44px 24px!important;
      background:rgba(255,255,255,.98)!important;
      border-left:1px solid var(--border)!important;
      border-top:0!important;
      transition:right .35s ease!important;
    }
    .nav-links.open { right:0!important; }
    .nav-links a { font-size:1.18rem!important; }
    .hamburger { display:flex!important; }
    .hero { padding:48px 0 60px!important; }
    .hero h1 { font-size:3.45rem!important; }
    .hero-photo-frame { width:min(100%,400px)!important; height:390px!important; border-radius:30px!important; }
    .hero-photo-frame img { border-radius:22px!important; }
    .enablement-grid { grid-template-columns:1fr!important; }
  }

  @media(max-width:520px) {
    .hero h1 { font-size:3rem!important; }
    .hero-subtitle { font-size:.94rem!important; }
    .hero-photo-frame { height:340px!important; }
    .btn-solid,.btn-outline { width:100%!important; justify-content:center!important; }
  }
</style>
'''

if 'id="visual-refresh-v6-2026-08-07"' in t:
    t = re.sub(r'\n<style id="visual-refresh-v6-2026-08-07">.*?</style>\n', '\n'+css+'\n', t, count=1, flags=re.S)
else:
    t = t.replace('\n</head>', '\n'+css+'\n</head>', 1)

p.write_text(t, encoding='utf-8')

checks = [
    'Fisioterapia y rehabilitación',
    'Masajes y bienestar',
    'Servicio de fisioterapia habilitado',
    'Registro vigente en REPS',
    'hero-accent',
    'visual-refresh-v6-2026-08-07',
    'loadReviews();',
    'pasaporte.html',
    'APPS_SCRIPT_URL'
]
for s in checks:
    if s not in t:
        raise SystemExit('Falta validación: '+s)
print('V6 aplicada y validada')
