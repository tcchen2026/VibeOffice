// Gestures run through tests/lib/gestures.mjs in each app: each starting state, each key in the middle and each way
// out. After any of them no drag is left and the cursor is the pointer's own; a gesture that ended with the release
// changed what it drags and is one undo step that brings back exactly what was there before; any other ending changed
// nothing and recorded nothing (the pattern behind R1–R4, todos/UI_MOUSE_LOGIC_IMPROVE.md). Each starting state is
// made from scratch: a click on a shape already selected means something else (Lectern edits its text).
//
//   node --test --test-concurrency=1 tests/ui/gestures.test.mjs
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { open, pause } from '../lib/input.mjs';
import { cases, run, settled } from '../lib/gestures.mjs';

/** p is inside the page's viewport (a point off it hits nothing, and the case would test nothing) */
const onScreen = (t, p) => t.evaluate(`${p.x} >= 0 && ${p.y} >= 0 && ${p.x} < innerWidth && ${p.y} < innerHeight`);
/** run a gesture through every case on one tab per starting state, then through the brief cases in each mode from the
    first state; o: { open(mode), starts, modes: [{ name, on(t) }] (open runs on before it builds what is dragged),
    from(t), to(t, a), state(t), steps(t) }; each start: { name, run(t), check(t) → true when the app is in that state } */
function matrix(name, o) {
  test(name, async (s) => {
    const runs = o.starts.map((start) => ({ start })).concat((o.modes || []).map((mode) => ({ start: o.starts[0], mode })));
    for (const { start, mode } of runs) {
      const t = await o.open(mode);
      try {
        for (const c of cases([start], { brief: !!mode })) {
          await s.test((mode ? `in ${mode.name}, ` : '') + c.name, async () => {
            await start.run(t); await pause(150);
            assert.equal(await start.check(t), true, `setup: ${start.name}`);
            const before = await o.state(t), n0 = await o.steps(t);
            const a = await o.from(t), b = await o.to(t, a);
            assert.ok(await onScreen(t, a) && await onScreen(t, b), `setup: the drag (${JSON.stringify([a, b])}) is on screen`);
            await run(t, a, b, c);
            assert.deepEqual(await settled(t), { dragging: false, cursorHeld: false }, 'no drag left, no cursor held');
            const after = await o.state(t);
            if (!c.commits) {
              assert.deepEqual(after, before, 'nothing changed');
              assert.equal(await o.steps(t), n0, 'nothing recorded');
              return;
            }
            assert.notDeepEqual(after, before, 'the release changed it');
            assert.equal(await o.steps(t), n0 + 1, 'one undo step');
            await t.evaluate(`L.ui.exec('undo')`); await pause(200);
            assert.deepEqual(await o.state(t), before, 'and one Undo brings back what was there');
          });
        }
        await s.test(`${mode ? `in ${mode.name}, ` : ''}from ${start.name}: no browser errors, no external requests`, () => t.clean());
      } finally { await t.close(); }
    }
  });
}

/* ---------------------------------------------------------------- Quire: moving a floating shape */
const quire = async (mode) => {
  const t = await open('quire');
  await t.evaluate(`(async () => {
    const { D } = L, d = D.newDoc();
    d.main.blocks = ['First paragraph of ordinary text for clicking.', 'Second paragraph of ordinary text.', 'Third paragraph.'].map((s) => D.para([D.text(s)]));
    const it = L.drawing.newShape('rect', 90, 60); L.drawing.setWrap(it, 'none'); it.name = 'Rectangle 1';
    it.float.posH = { rel: 'page', off: 330 }; it.float.posV = { rel: 'page', off: 200 }; d.main.blocks[2].runs.unshift(it);
    L.app.loadDoc(d, 'UI test');
    await new Promise((r) => setTimeout(r, 600));
  })()`);
  if (mode) { await mode.on(t); await pause(300); }
  return t;
};
const quireText = (t, b, o) => t.evaluate(`(() => { const m = L.layout.posToDom(L.D.pos(L.D.doc.main.blocks[${b}], ${o})), r = document.createRange();
  r.setStart(m.node, m.offset); r.collapse(true); const x = r.getClientRects()[0] || r.getBoundingClientRect(); return { x: x.left + 1, y: x.top + x.height / 2 }; })()`);
const flt = (t) => t.evaluate(`(() => { const r = document.querySelector('#pages .flt').getBoundingClientRect(); return { x: Math.round(r.left + r.width / 2), y: Math.round(r.top + r.height / 2), l: Math.round(r.left), t: Math.round(r.top) }; })()`);
matrix('Quire: moving a floating shape', {
  open: quire,
  starts: [
    { name: 'the caret in the text', run: async (t) => t.click(await quireText(t, 0, 5)), check: (t) => t.evaluate(`!L.ed.objSel && L.ed.collapsed()`) },
    { name: 'text selected', run: async (t) => { await t.click(await quireText(t, 0, 5)); await t.key('End', { shift: true }); }, check: (t) => t.evaluate(`!L.ed.objSel && !L.ed.collapsed()`) },
    { name: 'the shape selected', run: async (t) => { await t.click(await quireText(t, 1, 3)); await t.click(await flt(t)); }, check: (t) => t.evaluate(`!!L.ed.objSel`) },
  ],
  from: flt,
  to: (t, a) => ({ x: a.x + 80, y: a.y + 40 }),
  state: async (t) => ({ box: await flt(t), text: await t.evaluate(`L.D.doc.main.blocks.map((b) => L.D.ptext(b)).join('|')`) }),
  steps: (t) => t.evaluate('L.D.histState().undo.length'),
  modes: [
    { name: 'Track Changes', on: (t) => t.evaluate(`L.ui.exec('trackChanges')`) },   // the command: its toolbar comes at once
    { name: 'zoom 150%', on: (t) => t.evaluate(`L.layout.setZoom(1.5)`) },
  ],
});

/* ---------------------------------------------------------------- Ledger: moving a shape, the fill handle */
const cell = (t, r, c) => t.evaluate(`(() => { const w = document.querySelector('#gridwrap').getBoundingClientRect(), b = L.grid.cellRect(${r}, ${c});
  return { x: Math.round(w.left + b.x + b.w / 2), y: Math.round(w.top + b.y + b.h / 2), r: Math.round(w.left + b.x + b.w), b: Math.round(w.top + b.y + b.h) }; })()`);
const gobj = (t) => t.evaluate(`(() => { const r = document.querySelector('#glayer .gobj').getBoundingClientRect(); return { x: Math.round(r.left + r.width / 2), y: Math.round(r.top + r.height / 2), l: Math.round(r.left), t: Math.round(r.top) }; })()`);
/** A1 alone selected, no object */
const oneCell = (t) => t.evaluate(`(() => { const s = L.grid.sel(), r = s.ranges[0]; return !L.grid.selectedObject() && s.ranges.length === 1 && r.r1 === 0 && r.c1 === 0 && r.r2 === 0 && r.c2 === 0; })()`);
const ledgerSteps = (t) => t.evaluate('L.grid.wb.undo.stack.length');
const ledgerCells = (t) => t.evaluate(`[0, 1, 2, 3, 4].map((r) => L.editor.cellText(L.grid.sheet(), r, 0))`);
matrix('Ledger: moving a shape', {
  open: async (mode) => {
    const t = await open('ledger');
    if (mode) { await mode.on(t); await pause(300); }
    await t.evaluate(`L.ui.exec('drawRect')`);
    await t.drag(await cell(t, 4, 4), await cell(t, 9, 6), 6);
    return t;
  },
  starts: [
    { name: 'a cell', run: async (t) => t.click(await cell(t, 0, 0)), check: oneCell },
    { name: 'a range', run: async (t) => t.drag(await cell(t, 0, 0), await cell(t, 2, 1), 4), check: (t) => t.evaluate(`(() => { const r = L.grid.sel().ranges[0]; return !L.grid.selectedObject() && r.r1 === 0 && r.c1 === 0 && r.r2 === 2 && r.c2 === 1; })()`) },
    { name: 'the shape selected', run: async (t) => { await t.click(await cell(t, 0, 0)); await t.click(await gobj(t)); }, check: (t) => t.evaluate(`!!L.grid.selectedObject()`) },
  ],
  from: gobj,
  to: (t, a) => ({ x: a.x + 80, y: a.y + 40 }),
  state: async (t) => ({ box: await gobj(t), cells: await ledgerCells(t) }),
  steps: ledgerSteps,
  modes: [
    { name: 'frozen panes', on: async (t) => { await t.click(await cell(t, 2, 1)); await t.evaluate(`L.ui.exec('freezePanes')`); } },
    { name: 'the second sheet', on: (t) => t.evaluate(`L.app.activateSheet(1)`) },
    { name: 'zoom 150%', on: (t) => t.evaluate(`L.grid.setZoom(150)`) },
  ],
});
matrix('Ledger: the fill handle', {
  open: async (mode) => {
    const t = await open('ledger');
    if (mode) { await mode.on(t); await pause(300); }
    await t.click(await cell(t, 0, 0)); await t.type('1'); await t.key('Enter');
    return t;
  },
  starts: [{ name: 'the cell', run: async (t) => t.click(await cell(t, 0, 0)), check: oneCell }],
  from: async (t) => { const c = await cell(t, 0, 0); return { x: c.r - 1, y: c.b - 1 }; },
  to: async (t) => { const c = await cell(t, 3, 0); return { x: c.x, y: c.y }; },
  state: ledgerCells,
  steps: ledgerSteps,
  modes: [
    { name: 'frozen panes', on: async (t) => { await t.click(await cell(t, 2, 1)); await t.evaluate(`L.ui.exec('freezePanes')`); } },
    { name: 'the second sheet', on: (t) => t.evaluate(`L.app.activateSheet(1)`) },
    { name: 'zoom 150%', on: (t) => t.evaluate(`L.grid.setZoom(150)`) },
  ],
});

/* ---------------------------------------------------------------- Lectern: moving and resizing a shape */
const lectern = async (mode) => {
  const t = await open('lectern');
  if (mode) { await mode.on(t); await pause(300); }
  await t.evaluate(`(async () => {
    const E = L.ed, M = L.model, T = L.txt;
    if (L.te.active()) L.te.end();
    const slide = L.ed.slide();
    const r = M.newShape(L.pres, slide, 'rect', 200, 150, 120, 80); r.id = 'rectR';
    const b = M.newTextBox(L.pres, slide, 450, 300, 300, 40, true); b.id = 'boxB'; b.tx.ps = [{ lvl: 0, pp: {}, rs: [{ t: 'Some words' }] }];
    slide.shapes = [r, b];
    E.sel = []; E.render(); L.panes.left.render();
    await new Promise((res) => setTimeout(res, 400));
  })()`);
  return t;
};
const lbox = (t, id) => t.evaluate(`(() => { const r = L.ed.elOf('${id}').getBoundingClientRect(); return { x: Math.round(r.left + r.width / 2), y: Math.round(r.top + r.height / 2) }; })()`);
const lstate = (t) => t.evaluate(`L.ed.slide().shapes.map((s) => [s.id, Math.round(s.x), Math.round(s.y), Math.round(s.w), Math.round(s.h)])`);
const lsteps = (t) => t.evaluate('L.hist.undo.length');
/** a point of the slide's background on screen (where the slide and the slide area overlap, on no shape) */
const slideBlank = async (t) => {
  const p = await t.evaluate(`(() => { const h = L.ed.host.getBoundingClientRect(), v = L.ed.scroller.getBoundingClientRect();
    const l = Math.max(h.left, v.left, 0) + 6, r = Math.min(h.right, v.left + L.ed.scroller.clientWidth, innerWidth) - 6;
    const tp = Math.max(h.top, v.top, 0) + 6, b = Math.min(h.bottom, v.top + L.ed.scroller.clientHeight, innerHeight) - 6;
    for (let y = b; y > tp; y -= 12) for (let x = l; x < r; x += 12) { const e = document.elementFromPoint(x, y);
      if (e && L.ed.host.contains(e) && !e.closest('.shp, [data-h], .ed-overlay')) return { x: Math.round(x), y: Math.round(y) }; }
    return null; })()`);
  assert.ok(p, 'a point of the slide background on screen');
  return p;
};
const rectSelected = (t) => t.evaluate(`L.ed.sel.length === 1 && L.ed.sel[0] === 'rectR' && !L.te.active()`);
matrix('Lectern: moving a shape', {
  open: lectern,
  modes: [
    { name: 'a slide with a transition', on: (t) => t.evaluate(`L.ed.slide().trans = { type: 'fade', spd: 'slow', click: true }`) },
    { name: 'the second slide', on: (t) => t.evaluate(`(() => { L.ui.exec('newSlide'); L.ed.goto(1); })()`) },
    { name: 'the Outline tab', on: (t) => t.evaluate(`L.panes.left.setTab('outline')`) },
    { name: 'zoom 150%', on: (t) => t.evaluate(`L.ed.setZoom(1.5)`) },
  ],
  starts: [
    { name: 'nothing selected', run: async (t) => t.click(await slideBlank(t)), check: (t) => t.evaluate(`L.ed.sel.length === 0 && !L.te.active()`) },
    { name: 'the shape selected', run: async (t) => { await t.click(await slideBlank(t)); await t.click(await lbox(t, 'rectR')); }, check: rectSelected },
    { name: "another box's text being edited", run: async (t) => { await t.click(await slideBlank(t)); await t.dblclick(await lbox(t, 'boxB')); }, check: (t) => t.evaluate(`L.te.active() && L.te.target().sh.id === 'boxB'`) },
  ],
  from: (t) => lbox(t, 'rectR'),
  to: (t, a) => ({ x: a.x + 80, y: a.y + 40 }),
  state: lstate,
  steps: lsteps,
});
matrix('Lectern: resizing a shape', {
  open: lectern,
  modes: [
    { name: 'a slide with a transition', on: (t) => t.evaluate(`L.ed.slide().trans = { type: 'fade', spd: 'slow', click: true }`) },
    { name: 'the second slide', on: (t) => t.evaluate(`(() => { L.ui.exec('newSlide'); L.ed.goto(1); })()`) },
    { name: 'the Outline tab', on: (t) => t.evaluate(`L.panes.left.setTab('outline')`) },
    { name: 'zoom 150%', on: (t) => t.evaluate(`L.ed.setZoom(1.5)`) },
  ],
  starts: [{ name: 'the shape selected', run: async (t) => { await t.click(await slideBlank(t)); await t.click(await lbox(t, 'rectR')); }, check: rectSelected }],
  from: (t) => t.el('.ed-overlay [data-h="se"]'),
  to: (t, a) => ({ x: a.x + 60, y: a.y + 40 }),
  state: lstate,
  steps: lsteps,
});
