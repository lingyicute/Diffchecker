// Optional browser regression checks; the application and npm test stay dependency-free.
// Install Playwright outside this checkout and point PLAYWRIGHT_PACKAGE at that package.
// Then install Chromium and WebKit with the Playwright CLI and run this file with Node.
'use strict';
const assert = require('node:assert/strict');
const { readFile } = require('node:fs/promises');
const path = require('node:path');
const { pathToFileURL } = require('node:url');
const { setTimeout: delay } = require('node:timers/promises');
const { chromium, webkit } = require(process.env.PLAYWRIGHT_PACKAGE || 'playwright');

const url = tool => pathToFileURL(path.resolve(__dirname, '..', `${tool}-compare.html`)).href;

async function openTools(page) {
  await page.locator('#toolsBtn').click();
  assert.ok(await page.locator('#panel').evaluate(el => el.matches(':modal')));
  for (let i = 0; i < 18; i++) {
    await page.keyboard.press('Tab');
    assert.ok(await page.evaluate(() => document.activeElement === document.body || !!document.activeElement.closest('#panel')));
  }
}

async function closeTools(page) {
  await page.keyboard.press('Escape');
  await page.waitForFunction(() => !document.getElementById('panel').open);
  assert.equal(await page.locator('#toolsBtn').getAttribute('aria-expanded'), 'false');
  assert.ok(await page.locator('#toolsBtn').evaluate(el => el === document.activeElement));
}

async function checkKeyboardViewport(page, selector) {
  const rect = await page.locator(selector).evaluate(el => {
    // Desktop browser shells cannot summon a phone keyboard. Model the visual viewport
    // reported by a phone while keeping the browser's actual layout viewport unchanged.
    const vv = window.visualViewport;
    Object.defineProperties(vv, {
      height: { value: 280, configurable: true },
      offsetTop: { value: 65, configurable: true },
      scale: { value: 1, configurable: true },
    });
    vv.dispatchEvent(new Event('resize'));
    const { top, bottom } = el.getBoundingClientRect();
    delete vv.height; delete vv.offsetTop; delete vv.scale;
    vv.dispatchEvent(new Event('resize'));
    return { top, bottom };
  });
  assert.ok(rect.top >= 64 && rect.bottom <= 346, `obscured dialog: ${JSON.stringify(rect)}`);
}

async function download(page, format) {
  await page.locator('#exportBtn').click();
  await page.locator(`[data-fmt=${format}]`).click();
  const event = page.waitForEvent('download');
  await page.locator('#exportGo').click();
  const file = await event;
  return readFile(await file.path());
}

async function checkText(page) {
  await page.goto(url('text'));
  await openTools(page);
  await page.locator('#addRule').click();
  await page.locator('#rules input').focus();
  await checkKeyboardViewport(page, '#panel');
  await closeTools(page);
  await page.evaluate(() => { opts.realtime = true; });
  await page.locator('#exampleBtn').click();
  assert.ok(await page.locator('#result').evaluate(el => el.getBoundingClientRect().top < 250));
  const saved = JSON.parse(await download(page, 'diffchecker'));
  assert.equal(saved.type, 'text');
  assert.notEqual(saved.left, saved.right);

  const position = await page.evaluate(() => {
    const lines = Array.from({ length: 90 }, (_, i) => 'line ' + i);
    txtA.value = lines.join('\n'); lines[65] = 'changed'; txtB.value = lines.join('\n');
    run(true);
    const box = $('diffScroll');
    window.scrollTo(0, window.scrollY + box.getBoundingClientRect().top - 475);
    $('nextBtn').click();
    return $('focus-0').getBoundingClientRect().top;
  });
  assert.ok(position > 0 && position + 18 <= 667, `off-screen text change: ${position}`);
}

async function checkExcel(page) {
  await page.goto(url('excel'));
  await openTools(page);
  await closeTools(page);
  await page.locator('#exampleBtn').click();
  const header = page.locator('#detA [data-header]');
  await header.focus();
  assert.ok(await header.evaluate(el => {
    el.value = '2'; el.dispatchEvent(new Event('change'));
    return el.isConnected && document.activeElement === el;
  }));
  assert.equal(await page.locator('#prevA tr.hdr th').textContent(), '2');

  const sheet = value => ({ name: 'Orders', header: 0, rows: [['ID', 'Note'], ['1', value]] });
  const session = { type: 'excel', version: 1, left: sheet('old'), right: sheet('new'), leftName: 'before.csv', rightName: 'after.csv' };
  await page.locator('#fileInput').setInputFiles({ name: 'orders.diffchecker', mimeType: 'application/zip', buffer: Buffer.from(JSON.stringify(session)) });
  await page.waitForFunction(() => document.getElementById('nameA').textContent === 'before.csv');
  const csv = await download(page, 'csv');
  assert.match(csv.toString(), /ID,Note/);
  const xlsx = await download(page, 'xlsx');
  assert.equal(xlsx.readUInt32BE(0), 0x504b0304);
  await page.locator('#fileInput').setInputFiles({ name: 'export.xlsx', mimeType: 'application/octet-stream', buffer: xlsx });
  await page.waitForFunction(() => document.getElementById('nameA').textContent === 'export.xlsx');
  assert.match(await page.locator('#prevA').textContent(), /Note/);

  const position = await page.evaluate(() => {
    const head = 'ID,Site,Priority,Status,Owner,Note', data = Array.from({ length: 90 }, (_, i) => `${i},North warehouse,Standard service,Still in progress,Johanna Smith,original note`);
    setWorkbook('A', parseCsv([head, ...data].join('\n'), ','), 'table-before.csv');
    data[65] = data[65].replace('original note', 'new note');
    setWorkbook('B', parseCsv([head, ...data].join('\n'), ','), 'table-after.csv');
    run();
    const box = $('gridWrap');
    window.scrollTo(0, window.scrollY + box.getBoundingClientRect().top - 475);
    $('nextBtn').click();
    return { top: $('focus-0').getBoundingClientRect().top, left: box.scrollLeft };
  });
  assert.ok(position.top > 0 && position.top + 18 <= 667, `off-screen table change: ${position.top}`);
  assert.ok(position.left > 0, 'changed cell must be brought into the horizontal viewport');
}

async function touchScrollChromium(page) {
  const cdp = await page.context().newCDPSession(page);
  const viewport = page.locator('#viewport');
  const state = () => page.evaluate(() => ({ scroll: window.scrollY, pan: S.py }));
  const swipe = async () => {
    await viewport.scrollIntoViewIfNeeded();
    const b = await viewport.boundingBox();
    const point = { x: b.x + b.width / 2, y: b.y + b.height / 2, id: 1 };
    await cdp.send('Input.dispatchTouchEvent', { type: 'touchStart', touchPoints: [point] });
    for (let i = 1; i <= 8; i++) {
      await cdp.send('Input.dispatchTouchEvent', { type: 'touchMove', touchPoints: [{ ...point, y: point.y - 14 * i }] });
      await delay(24);
    }
    await cdp.send('Input.dispatchTouchEvent', { type: 'touchEnd', touchPoints: [] });
    await delay(300);
  };
  await viewport.scrollIntoViewIfNeeded();
  const initial = await state();
  await swipe();
  const scrolled = await state();
  assert.ok(scrolled.scroll - initial.scroll > 30, 'one finger should scroll the page');
  assert.equal(scrolled.pan, initial.pan);
  await page.locator('#touchPan').tap();
  await viewport.scrollIntoViewIfNeeded();
  const active = await state();
  await swipe();
  const moved = await state();
  assert.ok(Math.abs(moved.scroll - active.scroll) < 3, 'pan mode should not scroll the page');
  assert.ok(Math.abs(moved.pan - active.pan) > 35, 'pan mode should move the comparison');
  await page.locator('#touchPan').tap();
  await cdp.detach();
}

async function checkImage(page, engine) {
  await page.goto(url('image'));
  await page.locator('#exampleBtn').click();
  assert.equal(await page.getByRole('button', { name: 'Original image', exact: true }).count(), 1);
  assert.equal(await page.locator('#dropA').getByRole('button', { name: 'Remove' }).count(), 1);
  assert.match(await page.locator('#viewport').evaluate(el => getComputedStyle(el).touchAction), /pan-y/);
  if (engine === chromium) await touchScrollChromium(page);
  else {
    await page.locator('#touchPan').tap();
    assert.equal(await page.locator('#viewport').evaluate(el => getComputedStyle(el).touchAction), 'none');
    await page.locator('#touchPan').tap();
  }
  await page.locator('#alignBtn').click();
  await page.locator('#alDx').focus();
  await checkKeyboardViewport(page, '#alignModal');
  await page.locator('#alignCancel').click();
  await page.locator('[data-mode=split]').click();
  const canvases = await page.evaluate(() => {
    const probe = canvasUsable;
    try {
      canvasUsable = () => false;
      $('modeSeg').querySelector('[data-mode=highlight]').click();
      return stage.querySelectorAll('canvas').length;
    } finally { canvasUsable = probe; }
  });
  assert.equal(canvases, 2, 'failed mode switch must keep both split canvases');
  assert.equal(await page.locator('[data-mode=split]').getAttribute('aria-pressed'), 'true');
  const png = await download(page, 'png');
  assert.equal(png.readUInt32BE(0), 0x89504e47);
}

(async () => {
  for (const engine of [chromium, webkit]) {
    const browser = await engine.launch();
    try {
      const page = await browser.newPage({ viewport: { width: 320, height: 568 }, isMobile: true, hasTouch: true, reducedMotion: 'reduce', locale: 'en-US' });
      page.setDefaultTimeout(8000);
      const errors = [], outbound = [];
      page.on('pageerror', e => errors.push(e.message));
      page.on('request', r => { if (/^https?:/.test(r.url())) outbound.push(r.url()); });
      for (const tool of ['text', 'excel', 'image']) {
        await page.goto(url(tool));
        assert.ok(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth), `${tool} overflows 320px`);
      }
      await page.setViewportSize({ width: 360, height: 667 });
      await checkText(page);
      await checkExcel(page);
      await checkImage(page, engine);
      assert.deepEqual(errors, []);
      assert.deepEqual(outbound, []);
      console.log(`${engine.name()}: mobile layout, gestures and offline workflows passed`);
    } finally { await browser.close(); }
  }
})().catch(e => { console.error(e); process.exitCode = 1; });
