const puppeteer = require('puppeteer');
const path = require('path');

(async () => {
  const browser = await puppeteer.launch({ headless: 'new' });

  const pages = [
    {
      file: 'jornada-autopistas-carta.html',
      out:  'jornada-autopistas-CARTA.png',
      w: 816,   // 8.5in @ 96dpi
      h: 1056,  // 11in  @ 96dpi
    },
    {
      file: 'jornada-autopistas-oficio.html',
      out:  'jornada-autopistas-OFICIO.png',
      w: 816,   // 8.5in @ 96dpi
      h: 1344,  // 14in  @ 96dpi
    },
  ];

  for (const p of pages) {
    const page = await browser.newPage();
    await page.setViewport({ width: p.w, height: p.h, deviceScaleFactor: 2 });
    const url = 'file:///' + path.resolve(__dirname, p.file).replace(/\\/g, '/');
    await page.goto(url, { waitUntil: 'networkidle2', timeout: 30000 });
    // esperar a que cargue la imagen del QR
    await new Promise(r => setTimeout(r, 2000));
    await page.screenshot({ path: p.out, fullPage: false });
    console.log('Guardado:', p.out);
    await page.close();
  }

  await browser.close();
})();
