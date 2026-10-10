// Search commands (public/common/search.js), checked in each app and both layouts (menus and toolbars; the
// ribbon of Paper 2016): everyday words find the command a user means, first.
//
//   node --test tools/search.test.mjs   (needs Chromium with CDP at 127.0.0.1:9222, as tools/shot.mjs)
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import path from 'node:path';

const root = path.resolve(path.dirname(new URL(import.meta.url).pathname), '..');

/* query → the command id that must come first, in both layouts unless the key says 'bars:' or 'ribbon:' */
const FIRST = {
  quire: {
    bold: 'bold', image: 'insertPicture', margins: 'pageSetup', hyperlnk: 'hyperlink', 'word count': 'wordCount',
    graph: 'insertChart', synonym: 'thesaurus', settings: 'optionsDlg', strikethrough: 'strike', pdf: 'exportPDF',
    'track changes': 'trackChanges', 'page numbers': 'pageNumbers', uppercase: 'changeCase', 'find': 'find',
  },
  ledger: {
    bold: 'bold', image: 'insertPicture', margins: 'pageSetup', freeze: 'freezePanes', sum: 'autoSum',
    graph: 'insertChart', pdf: 'exportPDF', 'conditional': 'conditionalFormat', hyperlnk: 'hyperlink', sort: 'sortDlg',
  },
  lectern: {
    image: 'insertPicture', 'bars:present': 'showFromStart', 'ribbon:present': 'showFromStart', pdf: 'saveAs', 'new slide': 'newSlide',
    hyperlnk: 'hyperlink', graph: 'insertChart', 'speaker notes': 'speakerNotes', 'notes page': 'viewNotes', options: 'optionsDlg',
  },
};

/* first results with no menu, toolbar or ribbon place (reached through a dialog) */
const NOWHERE = { quire: ['bars:strikethrough'], ledger: [], lectern: [] };

const page = (cases) => `
const out = {};
for (const [k, want] of ${JSON.stringify(Object.entries(cases))}) {
  const [lay, q] = k.includes(':') ? k.split(':') : [null, k];
  for (const layout of lay ? [lay] : ['bars', 'ribbon']) {
    const r = L.ui.searchCommands(q, layout);
    out[layout + ':' + q] = { want, got: r[0] ? r[0].id : null, where: r[0] ? r[0].where : '', top: r.slice(0, 3).map((x) => x.label) };
  }
}
out.none = L.ui.searchCommands('zzqx qqzz', 'bars').length;
return JSON.stringify(out);
`;

for (const app of Object.keys(FIRST)) {
  test(`${app}: everyday words find the command`, () => {
    const out = execFileSync('node', [path.join(root, 'tools/shot.mjs'), app + '/', JSON.stringify([{ js: page(FIRST[app]) }])], { encoding: 'utf8' });
    const line = out.split('\n').find((l) => l.startsWith('js -> '));
    assert.ok(line, out);
    const r = JSON.parse(JSON.parse(line.slice(6)));
    assert.equal(r.none, 0, 'nonsense finds nothing');
    delete r.none;
    const wrong = Object.entries(r).filter(([, v]) => v.got !== v.want).map(([k, v]) => `${k}: wanted ${v.want}, got ${v.got} (${v.top.join(' | ')})`);
    assert.deepEqual(wrong, []);
    const nowhere = Object.entries(r).filter(([k, v]) => !v.where && !NOWHERE[app].includes(k)).map(([k]) => k);
    assert.deepEqual(nowhere, [], 'the first result says where it is');
    assert.match(out, /no errors, no external requests/);
  });
}
