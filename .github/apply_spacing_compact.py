from pathlib import Path
import re

p = Path('index.html')
t = p.read_text(encoding='utf-8')

css = r'''
<style id="spacing-compact-v1-2026-08-07">
  /* Compactación aprobada: solo ritmo, márgenes, padding, gaps y alturas. */
  .nav,
  .nav.scrolled {
    padding:14px 0 0!important;
  }
  .nav-inner {
    gap:12px!important;
  }
  .brand-logo,
  .nav.scrolled .brand-logo {
    width:255px!important;
  }
  .nav-links {
    min-height:60px!important;
    gap:clamp(32px,3vw,44px)!important;
  }

  .hero {
    padding:52px 0 60px!important;
  }
  .hero-grid {
    gap:56px!important;
  }
  .hero-content {
    padding:8px 0!important;
  }
  .hero-enabled-badge {
    margin-bottom:20px!important;
  }
  .hero h1 {
    margin:0 0 20px!important;
  }
  .hero-subtitle {
    margin:0 0 24px!important;
    line-height:1.70!important;
  }
  .hero-trust-note {
    margin-top:15px!important;
  }
  .hero-photo-wrap {
    padding:14px 0!important;
  }
  .hero-photo-frame {
    width:min(100%,420px)!important;
    max-width:420px!important;
    height:415px!important;
  }

  .enablement-strip {
    padding:0 0 50px!important;
  }
  .enablement-item {
    padding:18px 22px!important;
  }
  .enablement-legal {
    margin-top:10px!important;
  }

  .services,
  .experiences,
  .about,
  .reviews,
  .plans,
  .booking,
  .faq,
  .pasaporte-section,
  .contact {
    padding-top:72px!important;
    padding-bottom:72px!important;
  }

  .services-header,
  .experiences-header {
    margin-bottom:38px!important;
  }
  .services-header p,
  .experiences-header p {
    margin-top:14px!important;
    line-height:1.65!important;
  }

  .services-grid {
    gap:22px!important;
  }
  .service-card-img {
    height:205px!important;
  }
  .service-card-body {
    padding:22px!important;
  }
  .service-card h3 {
    margin-bottom:10px!important;
  }
  .service-card-body > p:not(.service-ideal) {
    line-height:1.60!important;
  }
  .service-price {
    margin-top:16px!important;
  }

  .experience-grid {
    gap:22px!important;
  }
  .experience-card {
    padding:22px!important;
  }
  .experience-card::before {
    height:175px!important;
    margin:-22px -22px 18px!important;
  }
  .experience-card h3 {
    margin-bottom:9px!important;
  }
  .experience-card > p:not(.ideal) {
    line-height:1.58!important;
  }

  .about-grid {
    gap:54px!important;
  }
  .about-content h2 {
    margin:10px 0 14px!important;
  }
  .about-body p {
    line-height:1.68!important;
  }
  .about-facts-v2 {
    gap:12px!important;
    margin:20px 0 4px!important;
  }
  .about-facts-v2 > div {
    padding:15px!important;
  }

  .reviews-header {
    margin-bottom:32px!important;
  }
  .reviews-overall {
    margin-bottom:24px!important;
  }
  .reviews-grid {
    gap:20px!important;
  }
  .review-card {
    padding:21px!important;
  }

  .footer {
    padding:42px 0 26px!important;
  }
  .footer-inner {
    gap:24px!important;
  }
  .footer::after {
    margin-top:16px!important;
    padding-top:14px!important;
  }

  @media (max-width:900px) {
    .hero-grid {
      gap:32px!important;
    }
    .hero-photo-wrap {
      padding:0!important;
    }
    .about-grid {
      gap:36px!important;
    }
  }

  @media (max-width:768px) {
    .nav,
    .nav.scrolled {
      padding:10px 0!important;
    }
    .nav-inner {
      gap:10px!important;
    }
    .brand-logo,
    .nav.scrolled .brand-logo {
      width:155px!important;
    }
    .hero {
      padding:36px 0 44px!important;
    }
    .hero-grid {
      gap:28px!important;
    }
    .hero-enabled-badge {
      margin-bottom:16px!important;
    }
    .hero h1 {
      margin-bottom:16px!important;
    }
    .hero-subtitle {
      margin-bottom:20px!important;
      line-height:1.62!important;
    }
    .hero-trust-note {
      margin-top:12px!important;
    }
    .hero-photo-frame {
      height:320px!important;
      width:min(100%,370px)!important;
      max-width:370px!important;
    }
    .enablement-strip {
      padding-bottom:38px!important;
    }
    .enablement-item {
      padding:16px 18px!important;
    }
    .services,
    .experiences,
    .about,
    .reviews,
    .plans,
    .booking,
    .faq,
    .pasaporte-section,
    .contact {
      padding-top:52px!important;
      padding-bottom:52px!important;
    }
    .services-header,
    .experiences-header,
    .reviews-header {
      margin-bottom:28px!important;
    }
    .services-grid,
    .experience-grid,
    .reviews-grid {
      gap:18px!important;
    }
    .service-card-img {
      height:210px!important;
    }
    .service-card-body,
    .experience-card,
    .review-card {
      padding:20px!important;
    }
    .experience-card::before {
      margin:-20px -20px 16px!important;
      height:165px!important;
    }
    .about-grid {
      gap:30px!important;
    }
    .footer {
      padding:36px 0 22px!important;
    }
  }

  @media (max-width:520px) {
    .hero-photo-frame {
      height:300px!important;
    }
    .services,
    .experiences,
    .about,
    .reviews,
    .plans,
    .booking,
    .faq,
    .pasaporte-section,
    .contact {
      padding-top:48px!important;
      padding-bottom:48px!important;
    }
  }
</style>
'''

if 'id="spacing-compact-v1-2026-08-07"' in t:
    t = re.sub(r'\n<style id="spacing-compact-v1-2026-08-07">.*?</style>\n', '\n'+css+'\n', t, count=1, flags=re.S)
else:
    if '</head>' not in t:
        raise SystemExit('No se encontró </head>')
    t = t.replace('\n</head>', '\n'+css+'\n</head>', 1)

# Validar que no se tocaron integraciones esenciales.
for marker in ['loadReviews();', 'APPS_SCRIPT_URL', 'pasaporte.html', 'id="agenda"', 'id="experiencias"', 'id="servicios"']:
    if marker not in t:
        raise SystemExit('Falta marcador crítico: '+marker)

p.write_text(t, encoding='utf-8')
print('Compactación aplicada correctamente')
