// Real app Save/Save As, snapshot and recovery hooks over CDP. No downloads or
// recent-file writes: intercept only the final saveFile handoff in our own tabs.
// Synthetic blank fixtures test plumbing; corpus runs test authored-file fidelity.
// node tools/ooxml/save-matrix.mjs OUTDIR [quire|ledger|lectern]
import fs from 'node:fs';
import path from 'node:path';
import { ensureServer, openPage } from './cdp.mjs';

const args = process.argv.slice(2), macroOption = args.indexOf('--macro-files');
const macroFiles = macroOption < 0 ? {} : JSON.parse(fs.readFileSync(args.splice(macroOption, 2)[1], 'utf8'));
const checkerOption = args.indexOf('--checker-only'), checkerOnly = checkerOption >= 0;
if (checkerOnly) args.splice(checkerOption, 1);
const [destination, only] = args;
if (!destination) throw new Error('Usage: save-matrix.mjs OUTDIR [quire|ledger|lectern]');
const out = path.resolve(destination);
fs.mkdirSync(out, { recursive: true });
// Explicit expectations, independent of the production format table.
const facts = {
  quire: [
    ['docx', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'],
    ['docm', 'application/vnd.ms-word.document.macroEnabled.12', 'application/vnd.ms-word.document.macroEnabled.main+xml'],
    ['dotx', 'application/vnd.openxmlformats-officedocument.wordprocessingml.template', 'application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml'],
    ['dotm', 'application/vnd.ms-word.template.macroEnabled.12', 'application/vnd.ms-word.template.macroEnabledTemplate.main+xml'],
  ],
  ledger: [
    ['xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'],
    ['xlsm', 'application/vnd.ms-excel.sheet.macroEnabled.12', 'application/vnd.ms-excel.sheet.macroEnabled.main+xml'],
    ['xltx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.template', 'application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml'],
    ['xltm', 'application/vnd.ms-excel.template.macroEnabled.12', 'application/vnd.ms-excel.template.macroEnabled.main+xml'],
  ],
  lectern: [
    ['pptx', 'application/vnd.openxmlformats-officedocument.presentationml.presentation', 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml'],
    ['pptm', 'application/vnd.ms-powerpoint.presentation.macroEnabled.12', 'application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml'],
    ['ppsx', 'application/vnd.openxmlformats-officedocument.presentationml.slideshow', 'application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml'],
    ['ppsm', 'application/vnd.ms-powerpoint.slideshow.macroEnabled.12', 'application/vnd.ms-powerpoint.slideshow.macroEnabled.main+xml'],
    ['potx', 'application/vnd.openxmlformats-officedocument.presentationml.template', 'application/vnd.openxmlformats-officedocument.presentationml.template.main+xml'],
    ['potm', 'application/vnd.ms-powerpoint.template.macroEnabled.12', 'application/vnd.ms-powerpoint.template.macroEnabled.main+xml'],
  ],
};
const report = path.join(out, 'results.jsonl');
fs.writeFileSync(report, '');
const stop = await ensureServer();
let failed = 0;
try {
  for (const [app, variants] of Object.entries(facts)) {
    if (only && app !== only) continue;
    const page = await openPage(app, { timeout: 45000, persistence: false });
    try {
      await page.evaluate(`(${install.toString()})(${JSON.stringify(app)})`);
      if (checkerOnly) await page.evaluate(`(async () => { const ext = ${JSON.stringify(variants[0][0])}; await __matrix.open(await __matrix.write(__matrix.model(), ext), 'checker.' + ext); })()`);
      for (const fact of checkerOnly ? [] : variants) {
        let result;
        try { result = await page.evaluate(`(${exercise.toString()})(${JSON.stringify(fact)})`); }
        catch (e) {
          result = { status: 'failed', error: e.message,
            context: await page.evaluate('({ checks: __matrix.checks, dialogs: Array.from(document.querySelectorAll("[role=dialog]")).map(e => e.textContent) })').catch(() => null) };
          failed++;
        }
        for (const [kind, data] of Object.entries(result.artifacts || {})) {
          const dir = path.join(out, app, kind); fs.mkdirSync(dir, { recursive: true });
          fs.writeFileSync(path.join(dir, 'variant.' + fact[0]), Buffer.from(data, 'base64'));
        }
        delete result.artifacts;
        const row = { app, variant: fact[0], ...result };
        fs.appendFileSync(report, JSON.stringify(row) + '\n');
        console.log(`${app} ${fact[0]}: ${result.status}, ${result.checks?.length || 0} checks`);
        if (result.status === 'failed') throw new Error(app + ' ' + fact[0] + ' matrix failed');
      }
      // Leave the real checker visible just long enough for its screenshot.
      await page.evaluate(`(${checker.toString()})()`);
      const shot = await page.send('Page.captureScreenshot', { format: 'png' });
      fs.writeFileSync(path.join(out, app + '-compatibility.png'), Buffer.from(shot.data, 'base64'));
      const checks = await page.evaluate(`__matrix.click('Compatibility Checker', 'Cancel'); __matrix.pending.then(() => __matrix.checks)`);
      const requests = [...new Set(page.requests)];
      const errors = page.errors.slice();
      if (errors.length) failed++;
      fs.appendFileSync(report, JSON.stringify({ app, scenario: 'checker', status: errors.length ? 'failed' : 'ok', checks, errors, requests }) + '\n');
      if (macroFiles[app]) {
        const file = macroFiles[app], data = fs.readFileSync(file).toString('base64');
        const result = await page.evaluate(`(${macroConversion.toString()})(${JSON.stringify({ data, name: path.basename(file), free: variants[0][0] })})`);
        for (const [name, bytes] of Object.entries(result.artifacts)) {
          const dir = path.join(out, app, 'macros'); fs.mkdirSync(dir, { recursive: true });
          fs.writeFileSync(path.join(dir, name), Buffer.from(bytes, 'base64'));
        }
        delete result.artifacts;
        fs.appendFileSync(report, JSON.stringify({ app, scenario: 'macro-conversion', source: file, ...result }) + '\n');
      }
    } catch (e) { failed++; fs.appendFileSync(report, JSON.stringify({ app, status: 'failed', error: e.message }) + '\n'); }
    finally { await page.close(); }
  }
} finally { stop(); }
console.log(report);
if (failed) process.exitCode = 1;

function install(app) {
  const L = window.L, h = window.__corpusHooks, A = L.app;
  const M = window.__matrix = { app, checks: [], saved: [], decision: 'saved' };
  M.current = () => h.current();
  M.model = () => app === 'ledger' ? h.current().wb : h.current();
  M.info = () => h.docInfo(h.current());
  M.check = (condition, label) => { if (!condition) throw new Error(label); M.checks.push(label); };
  M.dialog = title => document.querySelector('[role="dialog"][aria-label="' + title + '"]');
  M.click = (title, label) => {
    const button = Array.from(M.dialog(title)?.querySelectorAll('.dlg-foot button') || []).find(b => b.textContent.trim() === label);
    if (!button) throw new Error('Missing ' + title + ' / ' + label);
    button.click();
  };
  M.until = async (test, label) => {
    const deadline = Date.now() + 15000;
    while (!test()) { if (Date.now() > deadline) throw new Error('Timed out: ' + label); await new Promise(r => setTimeout(r, 20)); }
  };
  M.encode = async blob => {
    const a = new Uint8Array(await blob.arrayBuffer()); let s = '';
    for (let i = 0; i < a.length; i += 16384) s += String.fromCharCode(...a.subarray(i, i + 16384));
    return btoa(s);
  };
  M.write = (doc, ext) => app === 'ledger' ? L.xlsxWrite.write(doc, { type: ext }) : app === 'quire' ? L.docx.write(doc, { format: ext }) : L.pptx.write(doc, { format: ext });
  M.password = value => { if (app === 'ledger') h.current().password = value; else if (app === 'quire') M.model().openPassword = value; else M.model().password = value; };
  M.open = async (blob, name, opts, password) => {
    const work = h.open(new File([blob], name, { type: blob.type }), opts);
    if (password) { await M.until(() => M.dialog('Password'), 'password prompt'); M.dialog('Password').querySelector('input').value = password; M.click('Password', 'OK'); }
    await work;
    if (L.show?.active) await L.show.end();
  };
  VO.opened = VO.saved = () => {};
  L.saveFile = async (name, blob) => { M.saved.push({ name, blob }); return M.decision; };
  // Ledger's dialog starts this asynchronous call but returns when the dialog
  // closes. Capture its completion without replacing the real implementation.
  if (app === 'ledger') {
    const save = A.saveAs;
    A.saveAs = function (...args) { return M.exporting = save.apply(this, args); };
  }
  M.saveAs = async (ext, name, cancel = false) => {
    const work = app === 'ledger' ? L.dlg.saveAs() : A.saveAs();
    await M.until(() => M.dialog('Save As'), 'Save As');
    const dialog = M.dialog('Save As'), select = dialog.querySelector('select');
    M.check(select.value === (app === 'ledger' ? h.current().type : A.fileType), 'Save As defaults to current variant');
    dialog.querySelector('input[type="text"]').value = name;
    select.value = ext; select.dispatchEvent(new Event('change', { bubbles: true }));
    M.click('Save As', cancel ? 'Cancel' : 'Save');
    await work;
    if (app === 'ledger' && !cancel) await M.exporting;
  };
}

async function exercise([ext, mime, contentType]) {
  const M = window.__matrix, L = window.L, h = window.__corpusHooks;
  M.checks = []; M.saved = []; M.decision = 'saved'; M.password(undefined);
  const artifacts = {};
  const verify = async (blob, name, password) => {
    M.check(blob.type === mime.toLowerCase(), name + ' MIME');
    let bytes = new Uint8Array(await blob.arrayBuffer());
    if (password) {
      M.check(bytes[0] === 0xd0 && bytes[1] === 0xcf, name + ' encrypted');
      bytes = await L.officeCrypto.decrypt(bytes, password);
    }
    const pkg = await L.opc.open(bytes);
    M.check(pkg.type(pkg.main) === contentType, name + ' main content type');
    artifacts[name] = await M.encode(new Blob([bytes]));
  };
  const original = await M.write(M.model(), ext);
  await verify(original, 'original');
  await M.open(original, 'variant.' + ext);
  M.check(M.model().ooxmlFormat === ext, 'Open detects variant');
  M.check(!L.ui.dialogOpen(), 'No compatibility warning at open');
  await L.app.save();
  M.check(M.saved.length === 1 && M.saved[0].name === 'variant.' + ext, 'Save keeps filename');
  await verify(M.saved[0].blob, 'save');

  const identity = JSON.stringify(M.info());
  await M.saveAs(ext, 'cancelled', true);
  M.check(JSON.stringify(M.info()) === identity && M.saved.length === 1, 'Cancel Save As keeps identity and does not download');
  M.decision = 'declined';
  await M.saveAs(ext, 'declined');
  M.check(JSON.stringify(M.info()) === identity, 'Declined download keeps identity');
  M.decision = 'saved';
  await M.saveAs(ext, 'renamed.' + ext);
  M.check(M.saved.at(-1).name === 'renamed.' + ext, 'Save As uses extension once');
  M.check(M.info().name === 'renamed.' + ext, 'Successful Save As updates identity');
  await verify(M.saved.at(-1).blob, 'save-as');

  const info = M.info(), draftName = info.draftName || info.name;
  M.check(draftName === 'renamed.' + ext, 'Draft name keeps variant');
  const draft = await h.snapshot(h.current());
  await verify(draft, 'draft');
  await M.open(draft, draftName, { draft: true, name: info.name, saved: true });
  M.check(M.info().name === info.name && M.model().ooxmlFormat === ext, 'Recovery keeps filename and variant');
  await L.app.save();
  M.check(M.saved.at(-1).name === 'renamed.' + ext, 'Recovered Save keeps filename');
  await verify(M.saved.at(-1).blob, 'recovered');

  const password = 'matrix-test-password';
  M.password(password);
  const encrypted = await h.snapshot(h.current());
  await verify(encrypted, 'encrypted-draft', password);
  await M.open(encrypted, draftName, { draft: true, name: info.name, saved: true }, password);
  M.check(M.info().name === info.name, 'Encrypted recovery keeps filename');
  await L.app.save();
  await verify(M.saved.at(-1).blob, 'encrypted-recovered', password);
  M.password(undefined);
  return { status: 'ok', checks: M.checks, artifacts };
}

async function checker() {
  const M = window.__matrix, L = window.L, h = window.__corpusHooks;
  M.checks = []; M.saved = []; M.decision = 'saved';
  let doc = M.model();
  const loss = n => L.opc.loss(doc, { id: 'matrix:' + n, what: 'This test object cannot be represented after the edit.', where: 'Test object ' + n, action: 'conversion' });
  loss(1);
  const draft = await h.snapshot(h.current()), info = M.info();
  M.check(!L.ui.dialogOpen() && L.opc.pendingLosses(doc).length === 1, 'Draft is silent and does not acknowledge losses');
  await M.open(draft, info.draftName || info.name, { draft: true, name: info.name, saved: true, lossState: structuredClone(info.lossState) });
  doc = M.model();
  M.check(!L.ui.dialogOpen() && L.opc.pendingLosses(doc).length === 1, 'Recovery keeps pending losses without a dialog');
  let pending = L.app.save();
  await M.until(() => M.dialog('Compatibility Checker'), 'checker on Save');
  M.click('Compatibility Checker', 'Cancel'); await pending;
  M.check(M.saved.length === 0 && L.opc.pendingLosses(doc).length === 1, 'Cancel checker prevents download and acknowledgement');
  M.decision = 'declined'; pending = L.app.save();
  await M.until(() => M.dialog('Compatibility Checker'), 'checker after cancellation');
  M.click('Compatibility Checker', 'Continue'); await pending;
  M.check(L.opc.pendingLosses(doc).length === 1, 'Declined download does not acknowledge losses');
  M.decision = 'saved'; pending = L.app.save();
  await M.until(() => M.dialog('Compatibility Checker'), 'checker after declined download');
  M.click('Compatibility Checker', 'Continue'); await pending;
  M.check(L.opc.pendingLosses(doc).length === 0, 'Successful download acknowledges current entries');
  const count = M.saved.length;
  await L.app.save();
  M.check(M.saved.length === count + 1 && !L.ui.dialogOpen(), 'Repeat save is quiet');
  loss(2); M.pending = L.app.save();
  await M.until(() => M.dialog('Compatibility Checker'), 'new loss checker');
  M.check(M.dialog('Compatibility Checker').textContent.includes('Test object 2') && !M.dialog('Compatibility Checker').textContent.includes('Test object 1'), 'New loss warns without repeating acknowledged entries');
}

async function macroConversion(o) {
  const M = window.__matrix, L = window.L, h = window.__corpusHooks;
  M.checks = []; M.saved = []; M.decision = 'saved';
  const original = new Blob([Uint8Array.from(atob(o.data), c => c.charCodeAt(0))]);
  await M.open(original, o.name);
  const doc = M.model(), info = M.info();
  const draft = await h.snapshot(h.current());
  // The corpus may contain unrelated unsupported objects. This test isolates
  // the newly introduced macro conversion after acknowledging those entries.
  L.opc.acknowledge(doc, doc.losses || []);
  const macros = async blob => {
    const pkg = await L.opc.open(new Uint8Array(await blob.arrayBuffer()));
    return Object.fromEntries(pkg.names.filter(n => /vbaProject\.bin$/i.test(n)).map(n => [n, Array.from(pkg.bytes(n))]));
  };
  const before = await macros(original);
  M.check(Object.keys(before).length > 0, 'Authored fixture contains VBA');
  M.check(JSON.stringify(await macros(draft)) === JSON.stringify(before), 'Draft keeps original VBA bytes');
  let pending = M.saveAs(o.free, 'converted');
  await M.until(() => M.dialog('Compatibility Checker'), 'macro conversion warning');
  M.check(M.dialog('Compatibility Checker').textContent.includes('Macros cannot be saved'), 'Save As warns about macros before download');
  M.click('Compatibility Checker', 'Cancel'); await pending;
  M.check(M.saved.length === 0 && M.info().name === info.name, 'Cancelling macro conversion keeps original identity');
  await L.app.save();
  M.check(JSON.stringify(await macros(M.saved.at(-1).blob)) === JSON.stringify(before), 'Save after cancellation still preserves VBA');
  M.check(!L.ui.dialogOpen(), 'Cancelled conversion does not warn on macro-enabled Save');
  const artifacts = { ['original.' + doc.ooxmlFormat]: await M.encode(original), ['saved.' + doc.ooxmlFormat]: await M.encode(M.saved.at(-1).blob) };
  pending = M.saveAs(o.free, 'converted');
  await M.until(() => M.dialog('Compatibility Checker'), 'macro warning repeats after cancellation');
  M.click('Compatibility Checker', 'Continue'); await pending;
  const converted = M.saved.at(-1).blob;
  M.check(M.saved.at(-1).name === 'converted.' + o.free && Object.keys(await macros(converted)).length === 0, 'Approved macro-free Save As removes VBA');
  artifacts['converted.' + o.free] = await M.encode(converted);
  return { status: 'ok', checks: M.checks, artifacts };
}
