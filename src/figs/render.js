const { chromium } = require('playwright');
const path = require('path');
(async () => {
  const browser = await chromium.launch();
  const files = process.argv.slice(2);
  for (const f of files) {
    const page = await browser.newPage({ deviceScaleFactor: 2.5, viewport: { width: 1300, height: 900 } });
    await page.goto('file://' + path.resolve(f));
    await page.waitForTimeout(300);
    const el = await page.$('#fig');
    const out = f.replace(/\.html$/, '.png');
    await el.screenshot({ path: out });
    console.log('wrote', out);
    await page.close();
  }
  await browser.close();
})();
