// Dev helper: screenshots the web build with Playwright (uses the system Chromium).
// Usage: node scripts/shot.mjs <url> <out.png> [--dark] [--w=1440] [--h=900] [--eval="js"] [--wait=ms]
import { chromium } from 'playwright-core';

const [url, out, ...rest] = process.argv.slice(2);
const opts = Object.fromEntries(rest.map((a) => {
  const [k, ...v] = a.replace(/^--/, '').split('=');
  return [k, v.length ? v.join('=') : true];
}));
const browser = await chromium.launch({ executablePath: process.env.CHROME || '/opt/pw-browsers/chromium', args: ['--no-sandbox'] });
const page = await browser.newPage({ viewport: { width: Number(opts.w ?? 1440), height: Number(opts.h ?? 900) }, colorScheme: opts.dark ? 'dark' : 'light', deviceScaleFactor: Number(opts.dpr ?? 1) });
const logs = [];
page.on('console', (m) => logs.push(`[${m.type()}] ${m.text()}`));
page.on('pageerror', (e) => logs.push(`[pageerror] ${e.message}`));
await page.goto(url, { waitUntil: 'networkidle' });
await page.waitForTimeout(Number(opts.wait ?? 600));
if (opts.eval) {
  const r = await page.evaluate(opts.eval);
  if (r !== undefined) console.log('eval:', typeof r === 'string' ? r : JSON.stringify(r));
  await page.waitForTimeout(Number(opts.after ?? 600));
}
await page.screenshot({ path: out });
if (logs.length) console.log(logs.join('\n'));
await browser.close();
