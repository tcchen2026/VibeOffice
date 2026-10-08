// Reproduce the slide-6 diagnostics (six copies, plus the deck without slide 6 and slide 6 alone); automated success is not PowerPoint acceptance.
// node timing-repair.mjs OUTDIR [--source saved-sample.pptx] [--sdk oxval.dll] [--dotnet PATH]
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { ensureServer, openPage } from '../../ooxml/cdp.mjs';

const args = process.argv.slice(2);
const opt = (name, fallback) => { const i = args.indexOf(name); return i < 0 ? fallback : args.splice(i, 2)[1]; };
const source = opt('--source', null), sdk = opt('--sdk', null), dotnet = opt('--dotnet', 'dotnet');
if (args.length !== 1) throw new Error('Usage: timing-repair.mjs OUTDIR [--source FILE] [--sdk DLL] [--dotnet PATH]');
const out = path.resolve(args[0]);
if (fs.existsSync(out) && fs.readdirSync(out).length) throw new Error('Use an empty output directory; existing Office test copies are never replaced');
fs.mkdirSync(out, { recursive: true });
const stop = await ensureServer();
let page, result;
try {
  page = await openPage('lectern');
  result = await page.evaluate(`(${variants.toString()})(${JSON.stringify(source ? fs.readFileSync(source).toString('base64') : null)})`);
  assert.deepEqual(page.errors, [], 'Browser errors');
} finally { await page?.close(); stop(); }
for (const [name, bytes] of Object.entries(result.files)) fs.writeFileSync(path.join(out, name), Buffer.from(bytes, 'base64'));

const profile = fs.mkdtempSync(path.join(os.tmpdir(), 'vo-slide6-lo-'));
try {
  const converted = path.join(profile, 'converted'); fs.mkdirSync(converted);
  const r = spawnSync('soffice', ['-env:UserInstallation=file://' + profile, '--headless', '--convert-to', 'pptx', '--outdir', converted, path.join(out, '1-as-saved.pptx')], { encoding: 'utf8', timeout: 90000 });
  if (r.error || r.status || !fs.existsSync(path.join(converted, '1-as-saved.pptx'))) throw new Error('LibreOffice conversion failed: ' + (r.error || r.stderr || r.stdout));
  fs.copyFileSync(path.join(converted, '1-as-saved.pptx'), path.join(out, '2-resaved-by-libreoffice.pptx'));
} finally { fs.rmSync(profile, { recursive: true, force: true }); }

const manifest = { source: source || 'L.app.buildSample()', structural: result.structural, officeAcceptance: 'pending', files: [] };
let failed = false;
for (const name of fs.readdirSync(out).filter(n => n.endsWith('.pptx')).sort()) {
  const file = path.join(out, name);
  const r = spawnSync('python3', [fileURLToPath(new URL('../../ooxml/package.py', import.meta.url)), 'check', file], { encoding: 'utf8' });
  if (r.error || !r.stdout) throw new Error('Package checker failed: ' + (r.error || r.stderr));
  const entry = { name, sha256: createHash('sha256').update(fs.readFileSync(file)).digest('hex'), package: JSON.parse(r.stdout), powerpointRepair: null };
  failed ||= entry.package.status !== 'ok' && !name.startsWith('2-');   // LibreOffice's resave is a control
  if (sdk) {
    const v = spawnSync(dotnet, [sdk, file], { encoding: 'utf8', timeout: 120000, maxBuffer: 10 * 1024 * 1024 });
    if (v.error || !v.stdout) throw new Error('SDK validator failed: ' + (v.error || v.stderr));
    entry.sdk = JSON.parse(v.stdout);
    failed ||= entry.sdk.status !== 'ok';
  }
  manifest.files.push(entry);
}
fs.writeFileSync(path.join(out, 'manifest.json'), JSON.stringify(manifest, null, 2) + '\n');
console.log(JSON.stringify({ out, files: manifest.files.length, automated: failed ? 'failed' : 'ok', officeAcceptance: 'pending' }));
if (failed) process.exitCode = 1;

async function variants(encoded) {
  const base = encoded ? Uint8Array.from(atob(encoded), c => c.charCodeAt(0)) : new Uint8Array(await (await L.pptx.write(L.app.buildSample())).arrayBuffer());
  const zip = await L.zip.read(base), name = 'ppt/slides/slide6.xml';
  const entry = zip.get(name);
  if (!entry) throw new Error('No slide 6 in the sample');
  const xml = await entry.text(), P = 'http://schemas.openxmlformats.org/presentationml/2006/main', A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const parse = () => new DOMParser().parseFromString(xml, 'application/xml');
  const all = (root, ns, local) => Array.from(root.getElementsByTagNameNS(ns, local));
  const original = parse();
  const effects = all(original, P, 'cTn').filter(n => n.hasAttribute('presetID'));
  const builds = all(original, P, 'bldP');
  const pair = n => n.getAttribute('spid') + ':' + n.getAttribute('grpId');
  const demand = (condition, reason) => { if (!condition) throw new Error(reason); };
  demand(effects.length === 5 && builds.length === 3, 'Expected sample: three paragraph entrances, zoom, teeter; three builds');
  demand(new Set(builds.map(pair)).size === builds.length, 'Duplicate (spid, grpId) build pair');
  for (const b of builds) {
    const shape = all(original, P, 'sp').find(s => all(s, P, 'cNvPr').some(n => n.getAttribute('id') === b.getAttribute('spid')));
    demand(shape && all(shape, A, 't').length, 'bldP must target a shape with text');
  }
  demand(new Set(builds.map(n => n.getAttribute('spid'))).size === 2, 'Expected repeated target with distinct build groups');
  const encode = bytes => { let s = ''; for (let i = 0; i < bytes.length; i += 32768) s += String.fromCharCode(...bytes.subarray(i, i + 32768)); return btoa(s); };
  const files = { '1-as-saved.pptx': encode(base) };
  const entries = [];
  for (const item of new Set(zip.values())) entries.push({ name: item.name, data: await item.bytes() });
  async function emit(file, mutate) {
    const dom = parse(); mutate(dom);
    const changed = new XMLSerializer().serializeToString(dom);
    const blob = await L.zip.write(entries.map(e => e.name === name ? { name, data: changed } : e));
    files[file] = encode(new Uint8Array(await blob.arrayBuffer()));
  }
  await emit('3-slide6-no-animations.pptx', dom => { for (const n of all(dom, P, 'timing')) n.remove(); });
  await emit('4-slide6-no-build-list.pptx', dom => { for (const n of all(dom, P, 'bldLst')) n.remove(); });
  await emit('5-slide6-no-teeter.pptx', dom => {
    const teeter = all(dom, P, 'cTn').filter(n => n.getAttribute('presetClass') === 'emph' && n.getAttribute('presetID') === '32');
    demand(teeter.length === 1, 'Expected one teeter');
    const group = teeter[0].getAttribute('grpId');
    let wrapper = teeter[0].parentElement;
    // Remove the now-empty subgroup as well; an empty childTnLst is invalid.
    while (wrapper) {
      const list = wrapper.parentElement;
      wrapper.remove();
      wrapper = list?.localName === 'childTnLst' && !list.children.length && list.parentElement?.localName === 'cTn' ? list.parentElement.parentElement : null;
    }
    for (const n of all(dom, P, 'bldP')) if (n.getAttribute('grpId') === group) n.remove();
  });
  await emit('6-slide6-ids-in-order.pptx', dom => {
    const map = new Map();
    all(dom, P, 'cTn').forEach((n, i) => { map.set(n.getAttribute('id'), String(i + 1)); n.setAttribute('id', String(i + 1)); });
    for (const n of all(dom, P, 'tn')) { const old = n.getAttribute('val'); demand(map.has(old), 'Unresolved timing target'); n.setAttribute('val', map.get(old)); }
  });
  /* broader probes, from the sample's model: is slide 6 the cause at all? */
  if (!encoded) for (const [file, keep] of [['7-without-slide6.pptx', (_, i) => i !== 5], ['8-only-slide6.pptx', (_, i) => i === 5]]) {
    const pres = L.app.buildSample(); pres.slides = pres.slides.filter(keep);
    files[file] = encode(new Uint8Array(await (await L.pptx.write(pres)).arrayBuffer()));
  }
  return { files, structural: { effects: effects.length, builds: builds.length, buildPairs: builds.map(pair) } };
}
