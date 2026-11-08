// The pages' Content-Security-Policy, kept in one place (docs/suite.md, Security). It goes into each page as a
// <meta http-equiv>, first in <head>, so it travels with the files: the site, the PWA's offline copy (served from
// the service worker's cache, where a server's header never reaches) and the desktop app's pages.
//
//   node tools/csp.mjs           write it into the pages and the sources that make pages (Ledger's template,
//                                tools/oss-licenses.py), then rebuild Ledger's page
//   node tools/csp.mjs --check   exit 1 if a page or source lacks the policy, has another, has it after a script,
//                                or has an inline script, an on…= handler or a javascript: URL
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';

const root = path.resolve(path.dirname(new URL(import.meta.url).pathname), '..');

/** only our own script files run; styles may be inline (style attributes everywhere); pictures, fonts and data
 * from the site itself or from inside the page (data:, blob:); connections to the site, the update check
 * (https://vibeoffice.work) and the desktop app's IPC (ipc: on macOS and Linux, https://ipc.localhost on Windows) */
export const POLICY = [
  "default-src 'self'",
  "script-src 'self'",
  "style-src 'self' 'unsafe-inline'",
  "img-src 'self' data: blob:",
  "font-src 'self' data:",
  "connect-src 'self' data: blob: https://vibeoffice.work ipc: https://ipc.localhost",
  "worker-src 'self'",
  "manifest-src 'self'",
  "frame-src 'none'",
  "object-src 'none'",
  "base-uri 'none'",
  "form-action 'none'",
].join('; ');

const META = `<meta http-equiv="Content-Security-Policy" content="${POLICY}">`;
const FOUND = /<meta http-equiv="Content-Security-Policy" content="[^"]*">/g;
/** the pages, and the sources other tools build pages from */
export const PAGES = ['public/index.html', 'public/lite.html', 'public/quire/index.html', 'public/lectern/index.html', 'public/ledger/index.html', 'public/oss-licenses.html',
  'public/updater/index.html'];
export const SOURCES = ['tools/ledger/build/template.html', 'tools/oss-licenses.py'];

/** what is wrong with a page's (or source's) text, if anything */
export function problems(text) {
  const out = [];
  const metas = text.match(FOUND) || [];
  if (metas.length !== 1) out.push(`${metas.length} policies`);
  else if (metas[0] !== META) out.push('another policy');
  const at = text.indexOf(metas[0]), script = text.search(/<script\b/);
  if (metas.length && script >= 0 && script < at) out.push('a script before the policy');
  /* data blocks (JSON, the Start Center's JSON-LD for search engines) never run */
  for (const m of text.matchAll(/<script\b([^>]*)>([\s\S]*?)<\/script>/g)) {
    if (/\bsrc=/.test(m[1]) ? m[2].trim() : !/type="application\/(?:ld\+)?json"/.test(m[1])) out.push(`an inline script: ${m[0].slice(0, 60)}`);
  }
  for (const m of text.matchAll(/<[a-z][^>]*\son[a-z]+\s*=/gi)) out.push(`a handler: ${m[0].slice(0, 60)}`);
  for (const m of text.matchAll(/\b(?:href|src|action)\s*=\s*["']?\s*javascript:/gi)) out.push(`a javascript: URL: ${m[0]}`);
  return out;
}

function write(file) {
  const abs = path.join(root, file);
  let text = fs.readFileSync(abs, 'utf8');
  if (text.match(FOUND)) text = text.replace(FOUND, META);
  else if (text.includes('<meta charset="utf-8">')) text = text.replace('<meta charset="utf-8">', `<meta charset="utf-8">\n${META}`);
  else text = META + '\n' + text;   // Ledger's template: the head after make.py's charset and viewport
  fs.writeFileSync(abs, text);
}

if (import.meta.url === `file://${process.argv[1]}`) {
  if (process.argv.includes('--check')) {
    let bad = 0;
    for (const file of [...PAGES, ...SOURCES]) {
      for (const p of problems(fs.readFileSync(path.join(root, file), 'utf8'))) { console.log(`${file}: ${p}`); bad++; }
    }
    console.log(bad ? `${bad} problem(s): node tools/csp.mjs writes the policy; scripts go in files` : 'the policy is in every page');
    process.exit(bad ? 1 : 0);
  }
  for (const file of [...SOURCES, ...PAGES.filter((p) => p !== 'public/ledger/index.html')]) write(file);
  execFileSync('python3', [path.join(root, 'tools/ledger/build/make.py')], { stdio: 'inherit' });
  console.log('written:', [...PAGES, ...SOURCES].join(', '));
}
