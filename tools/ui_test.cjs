// Teste de interface com Playwright: abre a plataforma, carrega o cronograma e a
// planilha de avanço e passa pelas 5 telas. As bibliotecas da CDN são servidas
// a partir de node_modules (rode `npm install` antes), então funciona sem internet.
// Uso: node tools/ui_test.cjs [planilha_de_avanco.xlsx]
const { chromium } = require('playwright');
const path = require('path');
const ROOT = path.resolve(__dirname, '..'), NM = path.join(ROOT, 'node_modules');
const AVANCO = path.resolve(process.argv[2] || path.join(ROOT, 'SE_CFII_Avanco_Obra.xlsx'));
const SHOTS = path.join(ROOT, 'test-output');
(async () => {
  const b = await chromium.launch({ args: ['--use-gl=swiftshader', '--enable-unsafe-swiftshader'] });
  const ctx = await b.newContext({ viewport: { width: 1500, height: 900 } });
  await ctx.route(/cdnjs\.cloudflare\.com.*xlsx\.full\.min\.js/, r => r.fulfill({ path: path.join(NM, 'xlsx/dist/xlsx.full.min.js'), contentType: 'application/javascript' }));
  await ctx.route(/cdn\.jsdelivr\.net\/npm\/three@0\.160\.0\/(.*)/, r => r.fulfill({ path: path.join(NM, 'three', r.request().url().split('three@0.160.0/')[1]), contentType: 'application/javascript' }));
  await ctx.route(/fonts\.(googleapis|gstatic)\.com/, r => r.fulfill({ body: '', contentType: 'text/css' }));
  const pg = await ctx.newPage(); const errs = [];
  pg.on('pageerror', e => errs.push('pageerror: ' + e.message));
  pg.on('console', m => { if (m.type() === 'error') errs.push('console: ' + m.text()); });
  const txt = async sel => (await pg.locator(sel).innerText()).replace(/\s+/g, ' ').slice(0, 300);
  const shot = n => pg.screenshot({ path: path.join(SHOTS, n + '.png') });
  await pg.goto('file://' + path.join(ROOT, 'SE_Campo_Formoso_II_3D.html')); await pg.waitForTimeout(1500);
  await pg.setInputFiles('#cr-file', path.join(ROOT, 'Cronograma_SE_Campo_Formoso_CurvaS.xlsx')); await pg.waitForTimeout(2000);
  console.log('1 CRONOGRAMA:', await txt('#view-cron')); await shot('1-cronograma');
  for (const [h, v, n] of [['curva-s', '#view-curva', '2-curva-s'], ['placar', '#view-placar', '3-placar'], ['eap', '#view-eap', '5-eap']]) {
    await pg.evaluate(h => location.hash = h, h); await pg.waitForTimeout(800);
    console.log(n.toUpperCase() + ':', await txt(v)); await shot(n);
  }
  await pg.evaluate(() => location.hash = 'acompanhamento-3d'); await pg.waitForTimeout(3000);
  await pg.setInputFiles('#f-xlsx', AVANCO); await pg.waitForTimeout(2500);
  console.log('4 3D:', await txt('#view-3d')); await shot('4-3d');
  console.log(errs.length ? 'ERROS:\n' + errs.join('\n') : 'Sem erros no console.');
  await b.close(); process.exit(errs.length ? 1 : 0);
})();
