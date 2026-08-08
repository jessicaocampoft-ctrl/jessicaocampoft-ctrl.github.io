from pathlib import Path
import re

p = Path('index.html')
t = p.read_text(encoding='utf-8')

about_pattern = re.compile(r'''  <!-- EQUIPO -->\n  <section class="about" id="equipo">.*?  </section>\n\n  <!-- RESEÑAS GOOGLE -->''', re.S)

about_new = '''  <!-- EQUIPO -->
  <section class="about" id="equipo">
    <div class="container">
      <div class="about-grid">
        <div class="about-photo-wrap fade-up">
          <img src="Imagen%20de%20perfil%201.jpg" alt="Cuidándote Fisioterapia" loading="lazy" decoding="async">
          <div class="about-badge about-badge-team">
            <span class="big">✓</span>
            <span class="small">Atención profesional</span>
          </div>
        </div>
        <div class="about-content fade-up">
          <div class="section-tag">Nuestro enfoque</div>
          <h2>Cuidamos tu movimiento con experiencia clínica y humana</h2>
          <div class="about-body" id="aboutBody">
            <p>Cuidándote Fisioterapia es un espacio dedicado a la valoración, el tratamiento y la rehabilitación del movimiento, con atención profesional, humana y personalizada.</p>
            <p>Nuestro equipo está conformado por fisioterapeutas comprometidos con acompañar a cada persona según sus necesidades, objetivos y proceso.</p>
            <p>Combinamos evaluación fisioterapéutica, ejercicio terapéutico, técnicas manuales y educación para que cada persona comprenda su proceso y participe activamente en él.</p>
            <p>Atendemos personas con dolor, lesiones, limitaciones funcionales o necesidades relacionadas con el movimiento, tanto en nuestra sede de Dosquebradas como a domicilio en Pereira y Dosquebradas.</p>
          </div>
          <div class="about-facts-v2 about-facts-team">
            <div><strong>Atención profesional</strong><span>Acompañamiento humano y personalizado.</span></div>
            <div><strong>Trabajo en equipo</strong><span>Procesos construidos desde las necesidades de cada persona.</span></div>
            <div><strong>Enfoque funcional</strong><span>Movimiento, ejercicio terapéutico y educación.</span></div>
          </div>
        </div>
      </div>
    </div>
  </section>

  <!-- RESEÑAS GOOGLE -->'''

t, n = about_pattern.subn(about_new, t, count=1)
if n != 1:
    raise SystemExit('No se encontró la sección Equipo esperada')

css = r'''
<style id="brand-team-header-2026-08-07">
  /* Desktop: la marca vuelve a quedar al lado del menú, conservando la firma. */
  @media (min-width:769px) {
    .nav,
    .nav.scrolled {
      padding:14px 0!important;
    }
    .nav-inner {
      flex-direction:row!important;
      align-items:center!important;
      justify-content:space-between!important;
      gap:42px!important;
      max-width:1200px!important;
    }
    .brand-logo,
    .nav.scrolled .brand-logo {
      width:250px!important;
      max-width:250px!important;
      flex:0 0 250px!important;
      align-items:center!important;
    }
    .brand-logo .logo-byline {
      display:block!important;
      margin-top:1px!important;
      font-family:var(--font-heading)!important;
      font-size:.9rem!important;
      font-style:italic!important;
      font-weight:500!important;
      line-height:1.15!important;
      color:#52b8ae!important;
      letter-spacing:.01em!important;
    }
    .nav-links {
      width:auto!important;
      min-height:auto!important;
      flex:1 1 auto!important;
      justify-content:flex-end!important;
      gap:clamp(24px,2.3vw,38px)!important;
      border-top:0!important;
    }
    .nav-links a {
      font-size:1.08rem!important;
    }
  }

  /* Sección institucional: Cuidándote como marca y equipo. */
  #equipo .section-tag {
    display:inline-flex!important;
    align-items:center!important;
    width:max-content!important;
    padding:7px 13px!important;
    border:1px solid rgba(92,201,190,.34)!important;
    border-radius:999px!important;
    background:#fff!important;
    color:var(--primary-dark)!important;
    font-family:var(--font-body)!important;
    font-size:.72rem!important;
    font-weight:700!important;
    letter-spacing:.08em!important;
    text-transform:uppercase!important;
    box-shadow:none!important;
  }
  #equipo .section-tag::before {
    content:''!important;
    display:inline-block!important;
    width:7px!important;
    height:7px!important;
    margin-right:8px!important;
    border-radius:50%!important;
    background:var(--primary)!important;
  }
  #equipo .about-content h2 {
    max-width:660px!important;
    font-family:var(--font-heading)!important;
    font-size:clamp(2.8rem,4.2vw,4rem)!important;
    line-height:1.02!important;
    font-weight:600!important;
    margin:16px 0 22px!important;
    color:var(--text)!important;
  }
  #equipo .about-body {
    max-width:660px!important;
  }
  #equipo .about-body p {
    margin:0 0 15px!important;
    line-height:1.68!important;
    color:var(--text-muted)!important;
  }
  #equipo .about-body p:last-child {
    margin-bottom:0!important;
  }
  #equipo .about-facts-team {
    margin-top:24px!important;
  }
  #equipo .about-badge-team {
    min-width:145px!important;
  }
  #equipo .about-badge-team .big {
    font-size:1.8rem!important;
    line-height:1!important;
    color:var(--primary)!important;
  }
  #equipo .about-badge-team .small {
    max-width:105px!important;
    line-height:1.25!important;
  }

  @media (max-width:768px) {
    #equipo .about-content {
      text-align:left!important;
    }
    #equipo .section-tag {
      margin:0!important;
    }
    #equipo .about-content h2 {
      font-size:clamp(2.5rem,11vw,3.35rem)!important;
      margin:14px 0 18px!important;
    }
    #equipo .about-body p {
      margin-bottom:13px!important;
      line-height:1.62!important;
    }
  }
</style>
'''

if 'id="brand-team-header-2026-08-07"' in t:
    t = re.sub(r'\n<style id="brand-team-header-2026-08-07">.*?</style>\n', '\n'+css+'\n', t, count=1, flags=re.S)
else:
    t = t.replace('\n</head>', '\n'+css+'\n</head>', 1)

checks = [
    '<span class="logo-byline">by Jessica Ocampo</span>',
    'Cuidamos tu movimiento con experiencia clínica y humana',
    'Nuestro equipo está conformado por fisioterapeutas',
    'about-badge-team',
    'brand-team-header-2026-08-07',
    'loadReviews();',
    'APPS_SCRIPT_URL',
    'pasaporte.html'
]
for s in checks:
    if s not in t:
        raise SystemExit('Falta validación: '+s)

# Jessica Ocampo solo puede permanecer como firma de la marca en esta sección de cambios.
if '<h2>Jessica Ocampo</h2>' in t:
    raise SystemExit('Sigue existiendo el título personal Jessica Ocampo')

p.write_text(t, encoding='utf-8')
print('Ajustes de marca/equipo aplicados y validados')
