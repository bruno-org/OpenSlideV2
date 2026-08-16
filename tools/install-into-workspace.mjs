#!/usr/bin/env node
/**
 * Publish this fork into a workspace, the reliable way.
 *
 * Copying the sources over node_modules looks like it works, but Vite keeps a
 * transform cache for files under node_modules and goes on serving the previous
 * version even after a restart. Packing with `npm pack` and reinstalling swaps
 * the whole package and invalidates that cache along with it.
 *
 * Usage: node tools/install-into-workspace.mjs [workspace-path]
 */

import { execSync } from 'node:child_process';
import { existsSync, readFileSync, rmSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const FORK_ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const CORE = join(FORK_ROOT, 'packages', 'core');
const workspace = process.argv[2] ? resolve(process.argv[2]) : resolve(FORK_ROOT, '..', 'Slides');
const { version } = JSON.parse(readFileSync(join(CORE, 'package.json'), 'utf8'));
const TGZ = `open-slide-core-${version}.tgz`;

const run = (cmd, cwd) => {
  console.log(`[run] ${cmd}`);
  execSync(cmd, { cwd, stdio: 'inherit' });
};

if (!existsSync(join(workspace, 'package.json'))) {
  console.error(`[error] no workspace found at ${workspace}`);
  process.exit(1);
}

// The deck fonts have to exist on the system, otherwise the editable PPTX opens
// with PowerPoint's automatic substitution and the letterforms change.
run('node tools/install-fonts.mjs', FORK_ROOT);
run('pnpm run build', CORE);
run(`npm pack --pack-destination "${FORK_ROOT}"`, CORE);

// Removing the folder first forces npm to unpack again even when the tarball
// keeps the same name every time.
rmSync(join(workspace, 'node_modules', '@open-slide', 'core'), { recursive: true, force: true });
run(`npm install "${join(FORK_ROOT, TGZ)}"`, workspace);

// Vite's cacheDir lives inside the package itself; this is left over from the
// previous install.
rmSync(join(workspace, 'node_modules', '@open-slide', 'core', 'node_modules', '.vite'), {
  recursive: true,
  force: true,
});

console.log('\n[ok] fork installed. Start the dev server with: npm run dev');
