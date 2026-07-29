#!/usr/bin/env node
/**
 * Install the fonts this project ships into the current user's system.
 *
 * Why it matters: the editable PPTX export names fonts by family. Whoever opens
 * the file without them installed gets PowerPoint's automatic substitution and
 * the letterforms change. With them installed, the file opens as it looks on
 * screen.
 *
 * No administrator rights needed: everything goes into the user profile (on
 * Windows, LOCALAPPDATA plus a registry entry under HKCU). Idempotent, so
 * running it again changes nothing.
 *
 * Usage: node tools/install-fonts.mjs [--list]
 */

import { execFileSync } from 'node:child_process';
import { copyFileSync, existsSync, mkdirSync, readdirSync, statSync } from 'node:fs';
import { homedir, platform } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const FORK_ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');

/** Families the runtime applies by default, in themes and in its own chrome. */
const FONT_DIRS = ['geist-sans', 'geist-mono'];

function findFontSource() {
  // pnpm does not hoist, so the package may sit in node_modules/geist or inside
  // the virtual store. Look in both.
  const direct = join(FORK_ROOT, 'node_modules', 'geist', 'dist', 'fonts');
  if (existsSync(direct)) return direct;

  const store = join(FORK_ROOT, 'node_modules', '.pnpm');
  if (existsSync(store)) {
    for (const entry of readdirSync(store)) {
      if (!entry.startsWith('geist@')) continue;
      const candidate = join(store, entry, 'node_modules', 'geist', 'dist', 'fonts');
      if (existsSync(candidate)) return candidate;
    }
  }
  return null;
}

function collectFonts(root) {
  const out = [];
  for (const dir of FONT_DIRS) {
    const full = join(root, dir);
    if (!existsSync(full)) continue;
    for (const file of readdirSync(full)) {
      if (!file.endsWith('.ttf')) continue;
      // Variable files register under a family name systems handle
      // inconsistently; the static faces give the exact per-weight control the
      // export maps onto.
      if (file.includes('[wght]') || file.includes('Variable')) continue;
      out.push(join(full, file));
    }
  }
  return out;
}

/** "Geist-SemiBoldItalic.ttf" -> "Geist SemiBold Italic" */
function faceName(file) {
  const base = file.replace(/\.ttf$/i, '');
  const [family, style] = base.split('-');
  const spaced = (style ?? '')
    .replace(/([a-z])([A-Z])/g, '$1 $2')
    .replace(/Ultra (\w)/, 'Ultra$1')
    .trim();
  const familyName = family.replace(/([a-z])([A-Z])/g, '$1 $2');
  return spaced && spaced !== 'Regular' ? `${familyName} ${spaced}` : familyName;
}

const source = findFontSource();
if (!source) {
  console.error('[error] package "geist" not found. Run pnpm install at the repository root.');
  process.exit(1);
}

const fonts = collectFonts(source);
if (fonts.length === 0) {
  console.error(`[error] no .ttf found under ${source}`);
  process.exit(1);
}

if (process.argv.includes('--list')) {
  for (const f of fonts) console.log(`${faceName(f.split(/[\\/]/).pop())}  <-  ${f}`);
  process.exit(0);
}

const os = platform();
let installed = 0;
let skipped = 0;

if (os === 'win32') {
  const target = join(
    process.env.LOCALAPPDATA ?? join(homedir(), 'AppData', 'Local'),
    'Microsoft',
    'Windows',
    'Fonts',
  );
  mkdirSync(target, { recursive: true });

  for (const src of fonts) {
    const file = src.split(/[\\/]/).pop();
    const dest = join(target, file);
    const name = `${faceName(file)} (TrueType)`;

    const upToDate = existsSync(dest) && statSync(dest).size === statSync(src).size;
    if (!upToDate) copyFileSync(src, dest);

    try {
      // Per-user font: the value holds the full path.
      execFileSync(
        'reg',
        [
          'add',
          'HKCU\\Software\\Microsoft\\Windows NT\\CurrentVersion\\Fonts',
          '/v',
          name,
          '/t',
          'REG_SZ',
          '/d',
          dest,
          '/f',
        ],
        { stdio: 'ignore' },
      );
      installed++;
    } catch {
      skipped++;
      console.warn(`[warn] could not register ${name}`);
    }
  }
  console.log(`[ok] installed ${installed} font(s) for the current user in ${target}`);
  if (skipped > 0) console.warn(`[warn] ${skipped} failed to register`);
  console.log('Restart PowerPoint for it to pick up the new fonts.');
} else if (os === 'darwin') {
  const target = join(homedir(), 'Library', 'Fonts');
  mkdirSync(target, { recursive: true });
  for (const src of fonts) {
    copyFileSync(src, join(target, src.split('/').pop()));
    installed++;
  }
  console.log(`[ok] copied ${installed} font(s) to ${target}`);
} else {
  const target = join(homedir(), '.local', 'share', 'fonts');
  mkdirSync(target, { recursive: true });
  for (const src of fonts) {
    copyFileSync(src, join(target, src.split('/').pop()));
    installed++;
  }
  try {
    execFileSync('fc-cache', ['-f'], { stdio: 'ignore' });
  } catch {
    console.warn('[warn] fc-cache not found; refresh the font cache manually.');
  }
  console.log(`[ok] copied ${installed} font(s) to ${target}`);
}
