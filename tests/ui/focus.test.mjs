// The keyboard back to the document (todos/UI_MOUSE_LOGIC_IMPROVE.md, Phase 3), in each app and in the Classic and
// Paper 2016 looks: after a dialog opened from a shortcut, a toolbar or ribbon button or the context menu and closed
// with Esc, Cancel or OK, after a menu command, and after a press on chrome with no control of its own, the element
// that had the keyboard has it again, with the same selection.
//
//   node --test --test-concurrency=1 tests/ui/focus.test.mjs
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { open, pause } from '../lib/input.mjs';

/** a point on a container that is not on any control in it (the container or a bare part of it is hit there) */
const blank = (t, selector) => t.evaluate(`(() => { const box = document.querySelector(${JSON.stringify(selector)}); if (!box) return null; const r = box.getBoundingClientRect();
  if (!r.width || !r.height) return null;
  for (let x = Math.floor(r.right) - 2; x > r.left; x -= 3) for (const y of [r.top + r.height / 2, r.top + 3, r.bottom - 3]) {
    const e = document.elementFromPoint(x, y); if (e && box.contains(e) && !e.closest('button, input, select, textarea, a, label, [tabindex], [contenteditable], .tb-btn, .tb-split, .combo, .stab, .mb-item, .rb-tab, svg')) return { x, y: Math.round(y) }; }
  return null; })()`);
/** click the button of a command, on the ribbon tab that holds it in Paper 2016 */
async function commandButton(t, cmd, tab) {
  if (await t.evaluate(`!!document.querySelector('.rb-tabs')`)) await t.click(await t.el(`.rb-tab[data-tab="${tab}"]`));
  const b = await t.button(cmd);
  assert.ok(b, `the ${cmd} button`);
  await t.click(b);
}
const cancelButton = (t) => t.evaluate(`(() => { const b = [...document.querySelectorAll('.dlg .dlg-foot .btn')].find((b) => /cancel/i.test(b.textContent)); const r = b.getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height / 2 }; })()`);
const okButton = (t) => t.evaluate(`(() => { const b = document.querySelector('.dlg .dlg-foot .btn.primary'); const r = b.getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height / 2 }; })()`);

const APPS = {
  quire: {
    setup: (t) => t.evaluate(`(async () => {
      const { D } = L, d = D.newDoc();
      d.main.blocks = ['First paragraph of text.', 'Second paragraph of ordinary text.', 'Third paragraph.'].map((s) => D.para([D.text(s)]));
      L.app.loadDoc(d, 'UI test');
      await new Promise((r) => setTimeout(r, 600));
    })()`),
    /** select "paragraph" in the second paragraph with the mouse */
    into: async (t) => {
      const at = (o) => t.evaluate(`(() => { const m = L.layout.posToDom(L.D.pos(L.D.doc.main.blocks[1], ${o})), r = document.createRange(); r.setStart(m.node, m.offset); r.collapse(true); const b = r.getClientRects()[0] || r.getBoundingClientRect(); return { x: b.left + 1, y: b.top + b.height / 2 }; })()`);
      await t.drag(await at(7), await at(16), 6);
      return at(10);
    },
    selection: (t) => t.evaluate(`(() => { L.ed.syncSel(); const [a, b] = L.D.order(L.D.doc, L.ed.sel.a, L.ed.sel.f); return [L.D.doc.main.blocks.indexOf(a.p), a.o, L.D.doc.main.blocks.indexOf(b.p), b.o]; })()`),
    shortcut: ['d', { ctrl: true }], ok: true,
    context: 'Font',
    button: ['hyperlink', 'insert'],
    viewToggle: 'Ruler',
  },
  ledger: {
    setup: async () => {},
    into: async (t) => {
      const p = await t.evaluate(`(() => { const w = document.querySelector('#gridwrap').getBoundingClientRect(), b = L.grid.cellRect(2, 2); return { x: Math.round(w.left + b.x + b.w / 2), y: Math.round(w.top + b.y + b.h / 2) }; })()`);
      await t.click(p);
      return p;
    },
    selection: (t) => t.evaluate(`[L.grid.sel().r, L.grid.sel().c]`),
    shortcut: ['1', { ctrl: true }], ok: true,
    context: 'Cells',
    button: ['hyperlink', 'insert'],
    viewToggle: 'Status Bar',
  },
  lectern: {
    setup: (t) => t.evaluate(`(async () => {
      const E = L.ed, M = L.model;
      if (L.te.active()) L.te.end();
      E.goto(0);
      const slide = L.pres.slides[0];
      const b = M.newTextBox(L.pres, slide, 100, 100, 500, 40, true); b.id = 'boxA'; b.tx.ps = [{ lvl: 0, pp: {}, rs: [{ t: 'Alpha beta gamma delta epsilon' }] }];
      slide.shapes = [b];
      E.sel = []; E.render(); L.panes.left.render();
      await new Promise((r) => setTimeout(r, 400));
    })()`),
    /** edit the box's text and select "beta" with the mouse */
    into: async (t) => {
      const at = (n) => t.evaluate(`(() => { const w = document.createTreeWalker(L.ed.elOf('boxA'), NodeFilter.SHOW_TEXT); let k = ${n}, x;
        while ((x = w.nextNode())) { if (!x.data.replace(/\\u200b/g, '').length) continue; if (k <= x.length) { const r = document.createRange(); r.setStart(x, k); r.collapse(true); const b = r.getClientRects()[0] || r.getBoundingClientRect(); return { x: b.left + 1, y: b.top + b.height / 2 }; } k -= x.length; } return null; })()`);
      if (!(await t.evaluate('L.te.active()'))) { await t.click(await at(2)); await pause(200); if (!(await t.evaluate('L.te.active()'))) { await t.click(await at(2)); await pause(200); } }
      await t.drag(await at(6), await at(10), 6);
      return at(8);
    },
    selection: (t) => t.evaluate(`String(getSelection())`),
    shortcut: ['k', { ctrl: true }], ok: false,   // OK with no address asks for one
    context: 'Font',
    button: ['hyperlink', 'insert'],
    viewToggle: 'Ruler',
  },
};

const CHROME = ['.cmdbars', '#statusbar', '#drawrow', '#fbar', '.rb-tabs', '.tb-grip'];

for (const [app, A] of Object.entries(APPS)) for (const look of ['classic', 'paper2016']) {
  test(`${app}, ${look}: the keyboard stays with the document`, async (s) => {
    const t = await open(app, { look });
    try {
      await A.setup(t);
      /** put the keyboard in the document, run fn, and check the same element has it with the same selection */
      const keeps = async (fn) => {
        const p = await A.into(t);
        const home = await t.focused(), sel = await A.selection(t);
        await fn(p);
        await pause(250);
        assert.equal(await t.evaluate('L.ui.dialogOpen()'), false, 'no dialog left open');
        assert.equal(await t.focused(), home, 'the keyboard is back where it was');
        assert.deepEqual(await A.selection(t), sel, 'with the same selection');
      };
      const closers = [['Esc', () => t.esc()], ['Cancel', async () => t.click(await cancelButton(t))]];
      if (A.ok) closers.push(['OK', async () => t.click(await okButton(t))]);
      for (const [how, close] of closers) {
        await s.test(`a dialog from a shortcut, closed with ${how}`, () => keeps(async () => {
          await t.key(...A.shortcut); await pause(300);
          assert.equal(await t.evaluate('L.ui.dialogOpen()'), true, 'the dialog is open');
          await close();
        }));
      }
      await s.test(`a dialog from the ${look === 'classic' ? 'toolbar' : 'ribbon'}, closed with Cancel`, () => keeps(async () => {
        await commandButton(t, ...A.button); await pause(300);
        assert.equal(await t.evaluate('L.ui.dialogOpen()'), true, 'the dialog is open');
        await t.click(await cancelButton(t));
      }));
      await s.test('a dialog from the context menu, closed with Esc', () => keeps(async (p) => {
        await t.rightClick(p); await t.pick(A.context); await pause(300);
        assert.equal(await t.evaluate('L.ui.dialogOpen()'), true, 'the dialog is open');
        await t.esc();
      }));
      if (look === 'classic') {
        await s.test(`View ▸ ${A.viewToggle} chosen twice from the menu bar`, () => keeps(async () => {
          await t.menu('View', A.viewToggle); await pause(200);
          await t.menu('View', A.viewToggle);
        }));
        await s.test('a menu opened from the menu bar and closed with Esc', () => keeps(async () => {
          await t.menu('Edit'); await pause(150); await t.esc();
        }));
        /* Toolbar Options (») of a toolbar cut short (Standard and Formatting sharing a row at 1280 px): the pop-up
           holds the controls that don't fit, Esc closes it and they go back */
        await s.test('Toolbar Options (») opened, then closed with Esc', () => keeps(async () => {
          await t.evaluate(`(L.app.opts.twoRows = false, L.ui.toolbarPrefs(L.app.opts))`); await pause(300);
          try {
            await t.click(await t.el('#tb-standard .tb-more'));
            assert.equal(await t.evaluate(`!!document.querySelector('.tb-pop [data-cmd="help"]')`), true, 'Help is in the pop-up');
            await t.esc();
            assert.equal(await t.evaluate(`!document.querySelector('.tb-pop') && !!document.querySelector('#tb-standard [data-cmd="help"]')`), true, 'closed, Help back on the toolbar');
          } finally { await t.evaluate(`(L.app.opts.twoRows = true, L.ui.toolbarPrefs(L.app.opts))`); }
        }));
      } else {
        await s.test('the ribbon\'s tabs clicked', () => keeps(async () => {
          await t.click(await t.el('.rb-tab[data-tab="insert"]')); await pause(150);
          await t.click(await t.el('.rb-tab[data-tab="home"]'));
        }));
      }
      for (const sel of CHROME) {
        const p = await blank(t, sel);
        if (!p) continue;
        await s.test(`a press on ${sel} where it has no control`, () => keeps(async () => { await t.click(await blank(t, sel)); }));
      }
      /* a dialog's underlined letter with Alt (ui.dialog): Zoom's &200% is chosen, then Esc */
      await s.test('a dialog\'s access key pressed with Alt, then Esc', () => keeps(async () => {
        await t.evaluate(`L.ui.exec('zoomDlg')`); await pause(300);
        await t.key('2', { alt: true }); await pause(150);
        assert.match(await t.evaluate(`(() => { const r = [...document.querySelectorAll('.dlg input[type=radio]')].find((x) => x.checked); return r ? r.parentElement.textContent : ''; })()`), /^200\s?%$/, 'Alt+2 chose 200%');
        await t.esc();
      }));
      /* the status bar's zoom slider (ui.zoomSlider) is chrome too: the zoom changes, the keyboard and selection stay */
      const zoomNow = () => t.evaluate(`+document.querySelector('#statusbar .zs-track').getAttribute('aria-valuenow')`);
      await s.test('the zoom slider\'s thumb dragged', () => keeps(async () => {
        const z0 = await zoomNow(), tr = await t.el('#statusbar .zs-track'), th = await t.el('#statusbar .zs-thumb');
        await t.drag({ x: th.x, y: tr.y }, { x: Math.min(tr.r - 2, th.x + 24), y: tr.y }, 6);
        assert.ok(await zoomNow() > z0, `zoomed in from ${z0}%`);
      }));
      await s.test('the zoom slider dragged, then Esc', () => keeps(async () => {
        const z0 = await zoomNow(), tr = await t.el('#statusbar .zs-track'), th = await t.el('#statusbar .zs-thumb'), to = { x: Math.max(tr.l + 2, th.x - 24), y: tr.y };
        await t.press({ x: th.x, y: tr.y }); await t.glide({ x: th.x, y: tr.y }, to, 6);
        await t.esc(); await t.release(to);
        assert.equal(await zoomNow(), z0, 'the zoom is back');
      }));
      await s.test('the zoom slider\'s + clicked', () => keeps(async () => {
        const z0 = await zoomNow();
        await t.click(await t.evaluate(`(() => { const r = document.querySelectorAll('#statusbar .zs-step')[1].getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height / 2 }; })()`));
        assert.ok(await zoomNow() > z0, `zoomed in from ${z0}%`);
      }));
      await s.test('no browser errors, no external requests', () => t.clean());
    } finally { await t.close(); }
  });
}

/* ---------------------------------------------------------------- the same in the modes an app can be in (Classic) */
/* Each mode: on(t) after the app's setup; into and selection when the keyboard is somewhere else in that mode;
   shortcut for a dialog that opens there; context: false where the context menu is the browser's own. */
const qPoint = (t, story, b, o) => t.evaluate(`(() => { const m = L.layout.posToDom(L.D.pos(${story}[${b}], ${o})), r = document.createRange(); r.setStart(m.node, m.offset); r.collapse(true); const x = r.getClientRects()[0] || r.getBoundingClientRect(); return { x: x.left + 1, y: x.top + x.height / 2 }; })()`);
const qSel = (t) => t.evaluate(`(() => { L.ed.syncSel(); const [a, b] = L.D.order(L.D.doc, L.ed.sel.a, L.ed.sel.f); return [a.p.id, a.o, b.p.id, b.o]; })()`);
/** the page's own selection, as text */
const domSel = (t) => t.evaluate(`String(getSelection())`);
const MODES = {
  quire: [
    { name: 'Track Changes', on: (t) => t.evaluate(`L.ui.exec('trackChanges')`) },
    { name: 'a header being edited', on: async (t) => {
      await t.evaluate(`(async () => { const d = L.D.doc; d.hf.h1 = { kind: 'hdr', id: 'h1', blocks: [L.D.para([L.D.text('Header text here')], { style: 'Header' })] }; d.sect.refs = { hdr: { default: 'h1' }, ftr: {} };
        L.app.loadDoc(d, 'UI test'); await new Promise((r) => setTimeout(r, 600)); L.app.editHeaderFooter('hdr', 0); await new Promise((r) => setTimeout(r, 400)); })()`);
    },
    into: async (t) => { const st = `L.D.doc.hf.h1.blocks`; await t.drag(await qPoint(t, st, 0, 7), await qPoint(t, st, 0, 11), 6); return qPoint(t, st, 0, 9); },
    selection: qSel },
    { name: 'Outline view', on: (t) => t.evaluate(`L.ui.exec('viewOutline')`), selection: qSel },
  ],
  ledger: [
    { name: 'frozen panes', on: (t) => t.evaluate(`(() => { L.grid.select(1, 1); L.ui.exec('freezePanes'); })()`) },
    { name: 'the second sheet', on: (t) => t.evaluate(`L.app.activateSheet(1)`) },
    { name: 'zoom 150%', on: (t) => t.evaluate(`L.grid.setZoom(150)`) },
  ],
  lectern: [
    { name: 'the speaker notes being edited', on: async () => {},
      into: async (t) => { const p = await t.el('.notes-in'); await t.click(p); await t.key('End'); await t.type(' some notes'); for (let i = 0; i < 5; i++) await t.key('ArrowLeft', { shift: true }); return p; },
      selection: domSel, shortcut: ['p', { ctrl: true }], context: false, ok: false },
    { name: 'the Slides pane holding the keyboard', on: async () => {},
      into: async (t) => { const p = await t.el('.th[data-i="0"] .th-frame'); await t.click(p); return p; },
      selection: (t) => t.evaluate(`[L.ed.idx, [...L.app.slideSel].join(), L.app.focusArea]`), shortcut: ['p', { ctrl: true }], context: false, ok: false },
    { name: 'a slide with a transition', on: (t) => t.evaluate(`(() => { L.ed.slide().trans = { type: 'fade', spd: 'slow', click: true }; L.panes.left.render(); })()`) },
  ],
};
for (const [app, A] of Object.entries(APPS)) for (const M of MODES[app]) {
  test(`${app}, in ${M.name}: the keyboard stays with the document`, async (s) => {
    const t = await open(app, { look: 'classic' });
    try {
      await A.setup(t);
      await M.on(t); await pause(300);
      const into = M.into || A.into, selection = M.selection || A.selection, shortcut = M.shortcut || A.shortcut;
      const keeps = async (fn) => {
        const p = await into(t);
        const home = await t.focused(), sel = await selection(t);
        await fn(p);
        await pause(250);
        assert.equal(await t.evaluate('L.ui.dialogOpen()'), false, 'no dialog left open');
        assert.equal(await t.focused(), home, 'the keyboard is back where it was');
        assert.deepEqual(await selection(t), sel, 'with the same selection');
      };
      const closers = [['Esc', () => t.esc()], ['Cancel', async () => t.click(await cancelButton(t))]];
      if (M.ok ?? A.ok) closers.push(['OK', async () => t.click(await okButton(t))]);
      for (const [how, close] of closers) {
        await s.test(`a dialog from a shortcut, closed with ${how}`, () => keeps(async () => {
          await t.key(...shortcut); await pause(300);
          assert.equal(await t.evaluate('L.ui.dialogOpen()'), true, 'the dialog is open');
          await close();
        }));
      }
      if (M.context !== false) {
        await s.test('a dialog from the context menu, closed with Esc', () => keeps(async (p) => {
          await t.rightClick(p); await t.pick(A.context); await pause(300);
          assert.equal(await t.evaluate('L.ui.dialogOpen()'), true, 'the dialog is open');
          await t.esc();
        }));
      }
      await s.test('a menu opened from the menu bar and closed with Esc', () => keeps(async () => { await t.menu('Edit'); await pause(150); await t.esc(); }));
      for (const sel of ['.cmdbars', '#statusbar']) {
        if (!(await blank(t, sel))) continue;
        await s.test(`a press on ${sel} where it has no control`, () => keeps(async () => { await t.click(await blank(t, sel)); }));
      }
      await s.test('no browser errors, no external requests', () => t.clean());
    } finally { await t.close(); }
  });
}
