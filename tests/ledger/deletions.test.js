/* What the user deletes (a sheet, a chart, an add-in) is recorded as deleted, not as the writer's loss alarm. */
const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs'), path = require('node:path'), os = require('node:os');
const L = require('./load.js');
L.layout = { invalidate() {}, touchMerges() {} };
require('../../public/ledger/js/ops.js');
const O = L.ops;
const corpus = process.env.VO_CORPORA || path.join(os.homedir(), 'corpora');
async function open(t, name) {
  const file = path.join(corpus, 'excel', name);
  if (!fs.existsSync(file)) { t.skip('Fetch pinned corpus: ' + file); return null; }
  return L.xlsxRead.read(fs.readFileSync(file));
}
async function losses(wb) {
  await L.xlsxWrite.bytes(wb);
  const where = prefix => (wb.losses || []).filter(e => e.id.startsWith(prefix)).map(e => e.where).sort();
  return { deleted: where('deleted:'), alarms: where('unreferenced:'), shown: (wb.losses || []).filter(e => e.notify && /^deleted:/.test(e.id)).length };
}

test('a deleted sheet takes its chart; a deleted chart takes its part; unedited saves record nothing', async t => {
  const name = 'openxml-sdk__81e530d2f928__basicspreadsheet.xlsx';
  let wb = await open(t, name); if (!wb) return;
  assert.deepEqual(await losses(wb), { deleted: [], alarms: [], shown: 0 });
  O.tx(wb, 'Delete sheet', () => O.deleteSheet(wb, wb.sheets[0]));
  assert.deepEqual(await losses(wb), { deleted: ['xl/charts/chart1.xml', 'xl/worksheets/sheet1.xml'], alarms: [], shown: 0 });
  wb = await open(t, name);
  const sh = wb.sheets[0];
  O.tx(wb, 'Delete chart', () => O.setDrawings(sh, sh.drawings.filter(d => d.kind !== 'chart')));
  assert.deepEqual(await losses(wb), { deleted: ['xl/charts/chart1.xml'], alarms: [], shown: 0 });
});

test('a deleted add-in takes its picture and web extension; a copy still using them keeps them', async t => {
  const name = 'openxml-sdk__03f833d57306__Youtube.xlsx';
  let wb = await open(t, name); if (!wb) return;
  const sh = wb.sheets[0];
  O.tx(wb, 'Delete add-in', () => O.setDrawings(sh, sh.drawings.slice(1)));
  assert.deepEqual(await losses(wb), { deleted: ['xl/media/image1.png', 'xl/webextensions/webextension1.xml'], alarms: [], shown: 0 });
  wb = await open(t, name);
  const kept = wb.sheets[0];
  O.tx(wb, 'Copy and delete', () => O.setDrawings(kept, [L.clone(kept.drawings[0])]));
  assert.deepEqual((await losses(wb)).deleted, []);
});
