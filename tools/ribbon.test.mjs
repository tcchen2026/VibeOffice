// The Paper 2016 ribbon of each app (public/<app>/js/ribbon.js, drawn by common/ribbon.js), checked in the page:
// every command id in a tab, the Quick Access Toolbar, File or a drop-down menu exists; every button
// drawn has an icon; every icon name exists; and every command of the classic menus is reachable
// from the ribbon, apart from the listed exceptions.
//
//   node --test tools/ribbon.test.mjs   (needs Chromium with CDP at 127.0.0.1:9222, as tools/shot.mjs)
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import path from 'node:path';

const root = path.resolve(path.dirname(new URL(import.meta.url).pathname), '..');

/* menu commands with no place in the ribbon, and why */
const EXCEPTIONS = {
  quire: {
    clearContents: 'the Delete key',
    customize: 'customizes the classic toolbars',
    editObject: 'double-click the object',
    gettingStarted: 'Help (F1)',
    keyboardHelp: 'Help (F1)',
    newDrawing: 'the same as Insert ▸ Shapes',
    showAssistant: 'Help (F1)',
  },
  ledger: {
    customize: 'customizes the classic toolbars',
    keyboardHelp: 'Help (F1)',
  },
  lectern: {
    clear: 'the Delete key',
    gettingStarted: 'Help (F1)',
  },
};

const CHECK = `
VO.look.set('paper2016');
await new Promise((r) => setTimeout(r, 300));
const ui = L.ui, s = L.ribbonSpec();
const tabs = s.tabs.concat(...(s.contextual || []).map((c) => c.tabs));
const ids = new Set(), missing = [], noIcon = [];
const add = (id, where) => { if (!ui.cmds[id]) missing.push(where + ':' + id); ids.add(id); };
const walk = (it, where) => {
  if (typeof it === 'string') { if (it !== '|') add(it, where); return; }
  if (it.row) return it.row.forEach((x) => walk(x, where));
  if (it.cmd) add(it.cmd, where);
  if (it.split) add(it.split, where);
};
const menuCmds = (items, out) => {
  for (const x of (typeof items === 'function' ? items() : items) || []) {
    if (typeof x === 'string') { if (x !== '-') out.add(x); }
    else if (x && x.cmd) out.add(x.cmd);
    else if (x && x.sub) menuCmds(x.sub, out);
  }
  return out;
};
/* the commands in the ribbon's drop-down menus: catch what each menu function would open */
const open0 = ui.openMenu;
ui.openMenu = (items) => { menuCmds(items, ids); return null; };
try {
  for (const t of tabs) for (const g of t.groups) {
    if (g.launcher) add(g.launcher, t.id);
    for (const it of g.items) { walk(it, t.id); for (const x of it.row || [it]) if (x && x.menu) try { x.menu({ left: 0, bottom: 0 }, document.body); } catch (e) {} }
  }
} finally { ui.openMenu = open0; }
ui.closeMenus();
(s.qat || []).forEach((x) => walk(x, 'qat'));
menuCmds(s.file(), new Set()).forEach((id) => add(id, 'file'));
for (const t of tabs) {
  L.ribbonUI.select(t.id);
  for (const b of document.querySelectorAll('.rb-panel .tb-btn')) if (!b.querySelector('svg')) noIcon.push(t.id + ':' + (b.dataset.cmd || b.textContent));
}
L.ribbonUI.select(s.tabs[0].id);
const badIcons = Object.entries(s.icons || {}).filter(([, v]) => !L.icons.raw[v]).map(([k, v]) => k + '→' + v);
const inMenus = new Set();
L.app.menuBar.menus.forEach((m) => menuCmds(m.items, inMenus));
const unreachable = [...inMenus].filter((id) => !ids.has(id) && !id.startsWith('tb_')).sort();
VO.look.set('classic');
return JSON.stringify({ missing, noIcon, badIcons, unreachable });
`;

for (const app of ['quire', 'ledger', 'lectern']) {
  test(`${app}: the ribbon's commands exist and cover the menus`, () => {
    const out = execFileSync('node', [path.join(root, 'tools/shot.mjs'), app + '/', JSON.stringify([{ js: CHECK, wait: 200 }])], { encoding: 'utf8' });
    const line = out.split('\n').find((l) => l.startsWith('js -> '));
    assert.ok(line, out);
    const r = JSON.parse(JSON.parse(line.slice(6)));
    assert.deepEqual(r.missing, [], 'command ids that do not exist');
    assert.deepEqual(r.badIcons, [], 'icon names that do not exist');
    assert.deepEqual(r.noIcon, [], 'buttons without an icon');
    assert.deepEqual(r.unreachable.filter((id) => !EXCEPTIONS[app][id]), [], 'menu commands missing from the ribbon');
    assert.deepEqual(Object.keys(EXCEPTIONS[app]).filter((id) => !r.unreachable.includes(id)), [], 'exceptions now in the ribbon (remove them from the list)');
    assert.match(out, /no errors, no external requests/);
  });
}
