#!/usr/bin/env node
/**
 * Fidelity check for the editable PPTX export.
 *
 * 1. Exports the same deck in both formats through the browser (its own Chrome,
 *    temporary profile, never touching a browser you have open).
 * 2. Pulls the per-page PNG out of the image PPTX. That is the reference: the
 *    faithful capture of the screen the image exporter already produces.
 * 3. Opens the editable PPTX in actual PowerPoint (through COM) and exports
 *    every slide.
 * 4. Compares them and reports the divergence per slide.
 *
 * Run from packages/core, with a dev server up:
 *   pnpm dev:demo                            # in another terminal, from the repo root
 *   node tools/verify-pptx-editable.mjs [url] [work-dir]
 *
 * Step 3 needs PowerPoint installed, so it only runs on Windows. Everything
 * else works anywhere.
 */

import { execFileSync } from 'node:child_process';
import { mkdirSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { chromium } from '@playwright/test';
import { unzipSync } from 'fflate';

const URL = process.argv[2] ?? 'http://localhost:5173/s/verify-pptx-export';
// Intermediate files (both .pptx and the PNGs) go to the system temp dir; pass
// a folder as the second argument to inspect them afterwards.
const WORK = process.argv[3] ?? join(tmpdir(), 'open-slide-pptx-check');
/** Per-channel difference above which a pixel counts as diverging. */
const CHANNEL_TOLERANCE = 24;
/**
 * Divergence ceiling per slide, in percent of pixels. What remains is edge
 * antialiasing and PowerPoint's own rasterisation; geometry (text position,
 * width and height) is checked separately and is the real proof of fidelity.
 *
 * With a font that ships the actual weight it sits around 2%. The worst known
 * case is a weight synthesised by the browser (800 in Arial, which only goes up
 * to bold): the stem comes out thinner in PowerPoint and the gap reaches 4.4%.
 */
const MAX_DIFF_PCT = 5;

rmSync(WORK, { recursive: true, force: true });
mkdirSync(WORK, { recursive: true });

const failures = [];
const check = (ok, msg) => {
  console.log(`${ok ? '  [ok]' : '  [FAIL]'} ${msg}`);
  if (!ok) failures.push(msg);
};

const browser = await chromium.launch({ channel: 'chrome', headless: true });
const page = await browser.newPage({ viewport: { width: 1600, height: 1000 } });
page.on('pageerror', (e) => console.log(`  [pageerror] ${e.message}`));

async function exportVia(menuLabel, outFile) {
  await page.locator('button[aria-label="Download"]').first().click();
  await page.waitForTimeout(400);
  const item = page.getByText(menuLabel, { exact: true }).first();
  await item.waitFor({ state: 'visible', timeout: 10_000 });
  const wait = page.waitForEvent('download', { timeout: 180_000 });
  await item.click();
  const download = await wait;
  const path = join(WORK, outFile);
  await download.saveAs(path);
  return path;
}

console.log('\n1. export the deck in both formats');
await page.goto(URL, { waitUntil: 'networkidle' });
await page.waitForTimeout(1200);

const refPptx = await exportVia('Export as image PPTX', 'reference.pptx');
check(!!refPptx, 'image PPTX exported (reference)');

const editPptx = await exportVia('Export as editable PPTX', 'editable.pptx');
check(!!editPptx, 'editable PPTX exported');

console.log('\n2. inspect the structure of the editable PPTX');
const zip = unzipSync(new Uint8Array(readFileSync(editPptx)));
const names = Object.keys(zip);
const slideFiles = names.filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n));
check(slideFiles.length > 0, `${slideFiles.length} slide(s) in the package`);

const decoder = new TextDecoder();
let totalRuns = 0;
for (const name of slideFiles) {
  const xml = decoder.decode(zip[name]);
  const runs = (xml.match(/<a:t>/g) ?? []).length;
  totalRuns += runs;
  // Well-formed XML is a hard requirement: PowerPoint rejects the whole file on
  // any invalid node.
  const parsed = checkXml(xml);
  check(parsed.ok, `${name}: well-formed XML${parsed.ok ? '' : ` (${parsed.error})`}`);
  check(runs > 0, `${name}: ${runs} editable text run(s)`);
}
check(
  totalRuns >= slideFiles.length,
  `deck has ${totalRuns} editable text runs in total (${slideFiles.length} slide(s))`,
);

// The reference comes from the image PPTX: the faithful capture of the page
// WITH its text. The images inside the editable one are the text-free
// background and serve only as a diagnostic aid.
const refZip = unzipSync(new Uint8Array(readFileSync(refPptx)));
for (const png of Object.keys(refZip).filter((n) => n.startsWith('ppt/media/'))) {
  writeFileSync(join(WORK, `ref-${png.split('/').pop()}`), Buffer.from(refZip[png]));
}
for (const png of Object.keys(zip).filter((n) => n.startsWith('ppt/media/'))) {
  writeFileSync(join(WORK, `bg-${png.split('/').pop()}`), Buffer.from(zip[png]));
}

console.log('\n3. render the editable PPTX in PowerPoint');
const psScript = join(WORK, 'render.ps1');
writeFileSync(
  psScript,
  `$ErrorActionPreference = 'Stop'
$app = New-Object -ComObject PowerPoint.Application
$pres = $app.Presentations.Open('${editPptx.replace(/'/g, "''")}', $true, $false, $false)
for ($i = 1; $i -le $pres.Slides.Count; $i++) {
  $out = Join-Path '${WORK.replace(/'/g, "''")}' ("ppt-slide$i.png")
  $pres.Slides.Item($i).Export($out, 'PNG', 1920, 1080)
}
$pres.Close()
$app.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
Write-Output 'render ok'
`,
  'utf8',
);

let rendered = false;
try {
  const out = execFileSync(
    'powershell.exe',
    ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', psScript],
    { encoding: 'utf8', timeout: 180_000 },
  );
  rendered = out.includes('render ok');
} catch (e) {
  console.log(`  [powerpoint error] ${String(e.message).slice(0, 400)}`);
}
check(rendered, 'PowerPoint opened the file and exported the slides');

if (rendered) {
  console.log('\n4. compare against the reference, pixel by pixel');
  const compare = await browser.newPage();
  await compare.goto('about:blank');

  const files = readdirSync(WORK);
  for (let i = 1; i <= slideFiles.length; i++) {
    const refFile = files.find((f) => f === `ref-image${i}.png`);
    const gotFile = files.find((f) => f === `ppt-slide${i}.png`);
    if (!refFile || !gotFile) {
      check(false, `slide ${i}: image pair missing`);
      continue;
    }
    const a = readFileSync(join(WORK, refFile)).toString('base64');
    const b = readFileSync(join(WORK, gotFile)).toString('base64');
    const bgFile = files.find((f) => f === `bg-image${i}.png`);
    const bgB64 = bgFile ? readFileSync(join(WORK, bgFile)).toString('base64') : null;
    const geo = bgB64 ? await measureGeometry(compare, bgB64, a, b) : null;
    if (geo) {
      // Horizontal placement is exact, so it gets a tight bound. Vertical has a
      // residue on large text whose CSS line-height is tighter than the font's
      // natural line: PowerPoint will not compress a line below that minimum
      // and nudges the text up by a few pixels. On a 1080px canvas that is
      // around 1%, and the width and height checks below still hold.
      check(
        Math.abs(geo.dx) <= 8,
        `slide ${i}: text horizontally in place (dx=${geo.dx.toFixed(1)}, limit 8px)`,
      );
      check(
        Math.abs(geo.dy) <= 16,
        `slide ${i}: text vertically in place (dy=${geo.dy.toFixed(1)}, limit 16px)`,
      );
      check(
        Math.abs(geo.widthPct) <= 3,
        `slide ${i}: text width preserved (${geo.widthPct.toFixed(1)}%, limit 3%)`,
      );
      check(
        Math.abs(geo.heightPct) <= 3,
        `slide ${i}: text height preserved (${geo.heightPct.toFixed(1)}%, limit 3%)`,
      );
    }

    const diff = await compare.evaluate(
      async ({ a, b, tol }) => {
        const load = (src) =>
          new Promise((res, rej) => {
            const img = new Image();
            img.onload = () => res(img);
            img.onerror = () => rej(new Error('decode failed'));
            img.src = src;
          });
        const [ia, ib] = await Promise.all([
          load(`data:image/png;base64,${a}`),
          load(`data:image/png;base64,${b}`),
        ]);
        const W = 1920;
        const H = 1080;
        const draw = (img) => {
          const c = document.createElement('canvas');
          c.width = W;
          c.height = H;
          const ctx = c.getContext('2d', { willReadFrequently: true });
          ctx.drawImage(img, 0, 0, W, H);
          return ctx.getImageData(0, 0, W, H).data;
        };
        const da = draw(ia);
        const db = draw(ib);
        let bad = 0;
        for (let p = 0; p < da.length; p += 4) {
          const d =
            Math.abs(da[p] - db[p]) +
            Math.abs(da[p + 1] - db[p + 1]) +
            Math.abs(da[p + 2] - db[p + 2]);
          if (d > tol * 3) bad++;
        }
        return { pct: (bad / (W * H)) * 100 };
      },
      { a, b, tol: CHANNEL_TOLERANCE },
    );
    // Deliberately loose ceiling: at weight 800 or 500 PowerPoint only has bold
    // on or off, so the stem comes out thinner than on screen. Geometry, above,
    // is what proves fidelity; this catches gross regression.
    check(
      diff.pct <= MAX_DIFF_PCT,
      `slide ${i}: ${diff.pct.toFixed(2)}% diverging pixels (ceiling ${MAX_DIFF_PCT}%)`,
    );
  }
}

await browser.close();
console.log(`\nfiles in ${WORK}`);
if (failures.length > 0) {
  console.error(`\n${failures.length} check(s) failed.`);
  process.exit(1);
}
console.log('all checks passed.');

/**
 * Compares only the text itself: the text-free background comes from inside the
 * editable PPTX, so the reference's text and the render's text can be isolated
 * and measured for offset and size, instead of just counting different pixels.
 */
async function measureGeometry(page, bgB64, refB64, gotB64) {
  return page.evaluate(
    async ({ bg, ref, got }) => {
      const load = (src) =>
        new Promise((res, rej) => {
          const img = new Image();
          img.onload = () => res(img);
          img.onerror = () => rej(new Error('decode'));
          img.src = src;
        });
      const [ibg, iref, igot] = await Promise.all([
        load(`data:image/png;base64,${bg}`),
        load(`data:image/png;base64,${ref}`),
        load(`data:image/png;base64,${got}`),
      ]);
      const W = 1920;
      const H = 1080;
      const draw = (img) => {
        const c = document.createElement('canvas');
        c.width = W;
        c.height = H;
        const ctx = c.getContext('2d', { willReadFrequently: true });
        ctx.drawImage(img, 0, 0, W, H);
        return ctx.getImageData(0, 0, W, H).data;
      };
      const dbg = draw(ibg);
      const TOL = 90;
      const stats = (data) => {
        let minX = W;
        let minY = H;
        let maxX = -1;
        let maxY = -1;
        let n = 0;
        let sx = 0;
        let sy = 0;
        for (let p = 0, k = 0; p < data.length; p += 4, k++) {
          const d =
            Math.abs(data[p] - dbg[p]) +
            Math.abs(data[p + 1] - dbg[p + 1]) +
            Math.abs(data[p + 2] - dbg[p + 2]);
          if (d <= TOL) continue;
          const x = k % W;
          const y = (k / W) | 0;
          n++;
          sx += x;
          sy += y;
          if (x < minX) minX = x;
          if (x > maxX) maxX = x;
          if (y < minY) minY = y;
          if (y > maxY) maxY = y;
        }
        return n === 0 ? null : { minX, minY, maxX, maxY, cx: sx / n, cy: sy / n };
      };
      const a = stats(draw(iref));
      const b = stats(draw(igot));
      if (!a || !b) return null;
      const wa = a.maxX - a.minX;
      const wb = b.maxX - b.minX;
      const ha = a.maxY - a.minY;
      const hb = b.maxY - b.minY;
      return {
        dx: b.cx - a.cx,
        dy: b.cy - a.cy,
        widthPct: (wb / wa - 1) * 100,
        heightPct: (hb / ha - 1) * 100,
      };
    },
    { bg: bgB64, ref: refB64, got: gotB64 },
  );
}

/** Dependency-free XML sanity check: looks for unbalanced tags. */
function checkXml(xml) {
  const stack = [];
  const re = /<\/?([A-Za-z_][\w.:-]*)([^>]*?)(\/?)>/g;
  let m = re.exec(xml);
  while (m) {
    const [full, name, , selfClose] = m;
    if (full.startsWith('<?') || full.startsWith('<!')) {
      m = re.exec(xml);
      continue;
    }
    if (full.startsWith('</')) {
      const open = stack.pop();
      if (open !== name) return { ok: false, error: `closed ${name} while expecting ${open}` };
    } else if (!selfClose) {
      stack.push(name);
    }
    m = re.exec(xml);
  }
  if (stack.length > 0) return { ok: false, error: `unclosed tags: ${stack.join(', ')}` };
  return { ok: true };
}
