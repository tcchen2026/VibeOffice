// Lectern with real mouse and keyboard input (tests/lib/input.mjs). A test named after an audit item (P1…) checks
// what PowerPoint does there and stays `todo` until the item is fixed (todos/UI_MOUSE_LOGIC_IMPROVE.md, Audit,
// Lectern); taking the mark off is part of the fix. Each test has its own tab, so a failure leaves nothing behind.
// The deck is built through the model (setup); everything the user does is real input.
//
//   node --test --test-concurrency=1 tests/ui/lectern.test.mjs
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { open, pause } from '../lib/input.mjs';

/** a tab on Lectern for one test */
async function lectern(fn) {
  const u = await open('lectern');
  try { await fn(u); } finally { await u.close(); }
}

/** a deck of `slides` slides (notes S1, S2, … name them) whose first slide holds `shapes`, JS using
    box(id, x, y, text, bold), rect(id, x, y, w, h, text) and table(id, x, y) (3×3, cells "r0c0"…); slide 1 shown */
const setup = (u, shapes, slides = 1) => u.evaluate(`(async () => {
  const E = L.ed, M = L.model, T = L.txt;
  while (L.pres.slides.length < ${slides}) L.ui.exec('newSlide');
  L.pres.slides.forEach((s, i) => { s.notes = 'S' + (i + 1); });
  if (L.te.active()) L.te.end();
  E.goto(0);
  const slide = L.pres.slides[0];
  const box = (id, x, y, text, b) => { const s = M.newTextBox(L.pres, slide, x, y, 500, 40, true); s.id = id; s.tx.ps = [{ lvl: 0, pp: {}, rs: [{ t: text, b: !!b }] }]; return s; };
  const rect = (id, x, y, w, h, text) => { const s = M.newShape(L.pres, slide, 'rect', x, y, w, h); s.id = id; if (text) s.tx = T.body([T.para(text, { algn: 'ctr' })], { anchor: 'ctr' }); return s; };
  const table = (id, x, y) => { const t = M.newTable(L.pres, slide, 3, 3, x, y, 600); t.id = id; t.tbl.rows.forEach((r, ri) => r.cells.forEach((c, ci) => { c.tx.ps = [T.para('r' + ri + 'c' + ci)]; })); return t; };
  slide.shapes = [${shapes}];
  E.sel = []; E.render(); L.panes.left.render();
  await new Promise((r) => setTimeout(r, 400));
})()`);

/** screen point just inside character n of a shape's text */
const at = (u, id, n) => u.evaluate(`(() => { const w = document.createTreeWalker(L.ed.elOf('${id}'), NodeFilter.SHOW_TEXT); let k = ${n}, t;
  while ((t = w.nextNode())) { if (!t.data.replace(/\\u200b/g, '').length) continue; if (k <= t.length) { const r = document.createRange(); r.setStart(t, k); r.collapse(true); const b = r.getClientRects()[0] || r.getBoundingClientRect(); return { x: b.left + 1, y: b.top + b.height / 2 }; } k -= t.length; } return null; })()`);
/** a shape's box on screen: centre x, y and edges l, t, r, b */
const shapeBox = (u, id) => u.evaluate(`(() => { const r = L.ed.elOf('${id}').getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height / 2, l: r.left, t: r.top, r: r.right, b: r.bottom }; })()`);
/** screen point of a slide point (pt) */
const slidePt = (u, x, y) => u.evaluate(`(() => { const r = L.ed.host.getBoundingClientRect(); return { x: r.left + ${x} * L.ed.scale, y: r.top + ${y} * L.ed.scale }; })()`);
const thumb = (u, i) => u.el(`.th[data-i="${i}"] .th-frame`);
/** a little right of the centre of a table cell (inside its text's line, past the text) */
const cell = (u, id, r, c) => u.evaluate(`(() => { const td = L.ed.elOf('${id}').querySelector('td[data-r="${r}"][data-c="${c}"]'); const b = td.getBoundingClientRect(); return { x: b.left + b.width / 2 + 10, y: b.top + b.height / 2 }; })()`);
/** what is edited (a shape id or null) and what is selected */
const state = (u) => u.evaluate(`({ editing: L.te.state ? L.te.state.id : null, sel: L.ed.sel.slice() })`);
const order = (u) => u.evaluate(`L.pres.slides.map((s) => s.notes).join(',')`);
const text = (u, id) => u.evaluate(`L.ed.shape('${id}') ? L.txt.paraText(L.ed.shape('${id}').tx.ps[0]) : null`);
const geo = (u, id) => u.evaluate(`(s => s && [s.x, s.y, s.w, s.h].map((v) => Math.round(v)))(L.ed.shape('${id}'))`);
const count = (u) => u.evaluate(`L.ed.slide().shapes.length`);
/** AutoShapes ▸ Lines ▸ tool on the Drawing toolbar */
async function lineTool(u, label) { await u.click(await u.el('[data-as]')); await u.pick('Lines', label); }
/** edit boxA ("Alpha beta gamma delta epsilon") with "beta" selected */
async function editBeta(u) {
  await u.click(await at(u, 'boxA', 3));
  await u.drag(await at(u, 'boxA', 6), await at(u, 'boxA', 10));
  assert.deepEqual(await u.evaluate('String(window.getSelection())'), 'beta', 'setup: "beta" selected');
}
const BOXES = `box('boxA', 60, 300, 'Alpha beta gamma delta epsilon', true), box('boxB', 60, 380, 'Second box with other words')`;
const RECT = `rect('rectR', 200, 150, 120, 80)`;
const OK_TEXT = 'Alpha Q gamma delta epsilon';

/* ---------- Phase 1: data loss and wrong edits ---------- */

test('P1 After a thumbnail click, a click on a shape takes the keyboard: Delete deletes the shape', () => lectern(async (u) => {
  await setup(u, RECT, 2);
  await u.click(await thumb(u, 0));
  await u.click(await shapeBox(u, 'rectR'));
  assert.deepEqual((await state(u)).sel, ['rectR'], 'setup: the rectangle is selected');
  await u.key('Delete');
  assert.equal(await order(u), 'S1,S2', 'no slide is deleted');
  assert.equal(await count(u), 0, 'the rectangle is deleted');
  u.clean();
}));

for (const [id, cmd, prop, value] of [['P2a', 'bold', 'b', true], ['P2b', 'alignCenter', 'algn', 'ctr']]) {
  test(`${id} ${cmd === 'bold' ? 'Bold' : 'Center'} on a dragged range of table cells formats those cells only`, () => lectern(async (u) => {
    await setup(u, `table('tbl', 100, 150)`);
    await u.click(await cell(u, 'tbl', 0, 0));
    await u.drag(await cell(u, 'tbl', 0, 0), await cell(u, 'tbl', 1, 1), 10);
    await u.click(await u.button(cmd));
    const got = await u.evaluate(`L.ed.shape('tbl').tbl.rows.map((r) => r.cells.map((c) => { const p = c.tx.ps[0]; return ${prop === 'b' ? '!!(p.rs[0] && p.rs[0].b)' : "(p.pp && p.pp.algn) || null"}; }))`);
    const off = prop === 'b' ? false : null;
    assert.deepEqual(got, [[value, value, off], [value, value, off], [off, off, off]], 'the four selected cells, not the rest');
    u.clean();
  }));
}

test('a range of table cells stays selected after formatting: a second command formats the same cells', () => lectern(async (u) => {
  await setup(u, `table('tbl', 100, 150)`);
  await u.click(await cell(u, 'tbl', 0, 0));
  await u.drag(await cell(u, 'tbl', 0, 0), await cell(u, 'tbl', 1, 1), 10);
  await u.click(await u.button('bold'));
  await u.click(await u.button('italic'));
  const got = await u.evaluate(`L.ed.shape('tbl').tbl.rows.map((r) => r.cells.map((c) => { const x = c.tx.ps[0].rs[0] || {}; return (x.b ? 'B' : '-') + (x.i ? 'I' : '-'); }))`);
  assert.deepEqual(got, [['BI', 'BI', '--'], ['BI', 'BI', '--'], ['--', '--', '--']]);
  u.clean();
}));

/* ---------- Phase 2: one drag lifecycle ---------- */

test('P5a Esc during a move cancels it: the shape stays where it was', () => lectern(async (u) => {
  await setup(u, RECT);
  const c = await shapeBox(u, 'rectR');
  await u.press(c); await u.glide(c, { x: c.x + 80, y: c.y + 40 }, 6);
  await u.esc();
  await u.glide({ x: c.x + 80, y: c.y + 40 }, { x: c.x + 100, y: c.y + 50 }, 2); await u.release({ x: c.x + 100, y: c.y + 50 });
  assert.deepEqual(await geo(u, 'rectR'), [200, 150, 120, 80], 'not moved');
  u.clean();
}));

test('P5b Esc during a resize cancels it: the shape keeps its size', () => lectern(async (u) => {
  await setup(u, RECT);
  const c = await shapeBox(u, 'rectR');
  await u.click(c);
  const corner = { x: c.r, y: c.b };   // the bottom-right sizing handle
  await u.press(corner); await u.glide(corner, { x: corner.x + 60, y: corner.y + 50 }, 6);
  await u.esc();
  await u.glide({ x: corner.x + 60, y: corner.y + 50 }, { x: corner.x + 70, y: corner.y + 60 }, 2); await u.release({ x: corner.x + 70, y: corner.y + 60 });
  assert.deepEqual(await geo(u, 'rectR'), [200, 150, 120, 80], 'not resized');
  u.clean();
}));

test('P5c Esc while dragging out a rectangle draws nothing', () => lectern(async (u) => {
  await setup(u, RECT);
  await u.click(await u.button('drawRect'));
  const a = await slidePt(u, 500, 100), b = await slidePt(u, 650, 200);
  await u.press(a); await u.glide(a, b, 6);
  await u.esc();
  await u.glide(b, { x: b.x + 10, y: b.y + 10 }, 2); await u.release({ x: b.x + 10, y: b.y + 10 });
  assert.equal(await count(u), 1, 'no rectangle added');
  assert.equal(await u.evaluate(`L.ed.scroller.classList.contains('drawing')`), false, 'Esc also ends the tool');
  u.clean();
}));

test('R1 a key pressed in the middle of a resize waits: the resize is one undo step, and two Undos bring back the original', () => lectern(async (u) => {
  await setup(u, RECT);
  const before = await geo(u, 'rectR');
  await u.click(await shapeBox(u, 'rectR'));
  const se = await u.el('.ed-overlay [data-h="se"]'), to = { x: se.x + 60, y: se.y + 40 };
  await u.press(se); await u.glide(se, to, 6);
  await u.key('ArrowRight');
  await u.release(to);
  const resized = await geo(u, 'rectR');
  assert.deepEqual(resized.slice(0, 2), before.slice(0, 2), 'the arrow key moved nothing');
  assert.ok(resized[2] > before[2] && resized[3] > before[3], 'the release resized it');
  await u.key('z', { ctrl: true });
  assert.deepEqual(await geo(u, 'rectR'), before, 'one Undo brings back the original');
  await u.key('z', { ctrl: true });
  assert.deepEqual(await geo(u, 'rectR'), before, 'and a second changes nothing');
  u.clean();
}));

test('Ctrl+Z in the middle of a resize undoes the resize, not the move before it; Ctrl+Y brings back nothing uncommitted', () => lectern(async (u) => {
  await setup(u, RECT);
  const c = await shapeBox(u, 'rectR');
  await u.drag(c, { x: c.x + 40, y: c.y }, 6);
  const moved = await geo(u, 'rectR');
  await u.click(await shapeBox(u, 'rectR'));
  const se = await u.el('.ed-overlay [data-h="se"]'), to = { x: se.x + 60, y: se.y + 40 };
  await u.press(se); await u.glide(se, to, 6);
  await u.key('z', { ctrl: true });
  await u.release(to);
  assert.deepEqual(await geo(u, 'rectR'), moved, 'the shape is as it was moved: the resize is undone, the move kept');
  await u.key('y', { ctrl: true });
  assert.deepEqual(await geo(u, 'rectR'), moved, 'nothing to redo');
  u.clean();
}));

test('P5d Esc during a scribble draws nothing, without an error', () => lectern(async (u) => {
  await setup(u, RECT);
  await lineTool(u, 'Scribble');
  const a = await slidePt(u, 400, 300), b = await slidePt(u, 520, 380);
  await u.press(a); await u.glide(a, b, 8);
  await u.esc();
  await u.glide(b, { x: b.x + 20, y: b.y - 30 }, 3); await u.release({ x: b.x + 20, y: b.y - 30 });
  assert.deepEqual(u.errors, [], 'no TypeError when the button is let go');
  assert.equal(await count(u), 1, 'no scribble added');
  u.clean();
}));

test('P5e Esc while dragging a thumbnail cancels the reorder', () => lectern(async (u) => {
  await setup(u, '', 3);
  const a = await thumb(u, 0), c = await thumb(u, 2);
  await u.press(a); await u.glide(a, { x: a.x, y: c.b + 10 }, 10);
  await u.esc();
  await u.glide({ x: a.x, y: c.b + 10 }, { x: a.x, y: c.b + 14 }, 2); await u.release({ x: a.x, y: c.b + 14 });
  assert.equal(await order(u), 'S1,S2,S3', 'the slides keep their order');
  u.clean();
}));

test('P5f A right-click during a move cancels it, as Esc does, and opens no menu', () => lectern(async (u) => {
  await setup(u, RECT);
  const c = await shapeBox(u, 'rectR'), b = { x: c.x + 70, y: c.y + 50 };
  await u.press(c); await u.glide(c, b, 6);
  await u.mouse('mousePressed', b, { button: 'right' }); await pause(60);
  await u.mouse('mouseReleased', b, { button: 'right' }); await pause(100);
  await u.glide(b, { x: b.x + 20, y: b.y + 20 }, 3); await u.release({ x: b.x + 20, y: b.y + 20 });
  assert.deepEqual(await geo(u, 'rectR'), [200, 150, 120, 80], 'not moved');
  assert.equal(await u.evaluate('L.ui.menuOpen()'), false, 'and opens no menu');
  u.clean();
}));

test('P10 Another drawing tool abandons an unfinished freeform: the next freeform has only its own points', () => lectern(async (u) => {
  await setup(u, '');
  await lineTool(u, 'Freeform');
  await u.click(await slidePt(u, 100, 100));
  await u.click(await slidePt(u, 200, 150));
  await u.click(await u.button('drawRect'));
  assert.equal(await u.evaluate(`document.querySelectorAll('.ed-overlay .draw-ghost').length`), 0, 'the unfinished freeform is gone from the slide');
  await u.drag(await slidePt(u, 600, 100), await slidePt(u, 700, 180));
  await lineTool(u, 'Freeform');
  await u.click(await slidePt(u, 100, 350));
  await u.click(await slidePt(u, 200, 420));
  await u.dblclick(await slidePt(u, 300, 380));
  const spans = await u.evaluate(`L.ed.slide().shapes.filter((s) => s.geom === 'custom').map((s) => [Math.round(s.y), Math.round(s.y + s.h)])`);
  assert.ok(spans.length >= 1, 'a freeform was drawn');
  for (const [top, bottom] of spans) assert.ok(bottom <= 160 || top >= 345, `a freeform spans y ${top}–${bottom}: it joins the abandoned points to the new ones`);
  u.clean();
}));

test('P14 A move whose release is lost (let go in another tab) does not follow the pointer afterwards', () => lectern(async (u) => {
  await setup(u, RECT);
  const c = await shapeBox(u, 'rectR'), b = { x: c.x + 70, y: c.y + 40 };
  await u.press(c); await u.glide(c, b, 6);
  await u.releaseInOtherTab();
  const before = await shapeBox(u, 'rectR');
  await u.glide(b, { x: b.x + 200, y: b.y + 120 }, 6);   // no button held
  const after = await shapeBox(u, 'rectR');
  assert.deepEqual([after.l, after.t], [before.l, before.t], 'the shape stays put while the pointer only hovers');
  u.clean();
}));

/* ---------- Phase 3: the keyboard back to the document ---------- */

for (const [id, how, show] of [
  ['P7a', 'Format ▸ Font…, Cancel', async (u) => { await u.menu('Format', 'Font'); return 'cancel'; }],
  ['P7b', 'Format ▸ Font…, Esc', async (u) => { await u.menu('Format', 'Font'); return 'esc'; }],
  ['P7c', 'the context menu\'s Font…, Cancel', async (u) => { await u.rightClick(await at(u, 'boxA', 8)); await u.pick('Font'); return 'cancel'; }],
  ['P7d', 'the Hyperlink button, Cancel', async (u) => { await u.click(await u.button('hyperlink')); return 'cancel'; }],
]) {
  test(`${id} After ${how}, the text selection is kept and typing replaces it`, () => lectern(async (u) => {
    await setup(u, BOXES);
    await editBeta(u);
    const close = await show(u);
    await pause(300);
    assert.equal(await u.evaluate('L.ui.dialogOpen()'), true, 'setup: the dialog is open');
    if (close === 'esc') await u.esc();
    else await u.click(await u.evaluate(`(() => { const b = [...document.querySelectorAll('.dlg .dlg-foot .btn')].find((b) => /cancel/i.test(b.textContent)); const r = b.getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height / 2 }; })()`));
    await u.type('Q');
    assert.equal(await text(u, 'boxA'), OK_TEXT, 'typing replaces "beta"');
    u.clean();
  }));
}

test('With the Slides pane holding the keyboard, Format ▸ Background… closed with Esc gives it back: Delete deletes the slide', () => lectern(async (u) => {
  await setup(u, '', 3);
  await u.click(await thumb(u, 0));
  await u.menu('Format', 'Background');
  assert.equal(await u.evaluate('L.ui.dialogOpen()'), true, 'setup: the dialog is open');
  await u.esc(); await pause(200);
  await u.key('Delete');
  assert.equal(await order(u), 'S2,S3', 'Delete deletes the slide');
  u.clean();
}));

test('Editing the speaker notes, Print (Ctrl+P) closed with Esc gives the keyboard back to the notes, the caret where it was', () => lectern(async (u) => {
  await setup(u, '');
  await u.click(await u.el('.notes-in'));
  await u.key('End');
  await u.type(' more');
  const notes = await u.evaluate('L.ed.slide().notes');
  assert.ok(notes.endsWith(' more'), 'setup: typed into the notes');
  await u.key('p', { ctrl: true }); await pause(300);
  assert.equal(await u.evaluate('L.ui.dialogOpen()'), true, 'setup: the Print dialog is open');
  await u.esc(); await pause(200);
  assert.equal(await u.focused(), 'notes-in', 'the notes have the keyboard');
  await u.type('X');
  assert.equal(await u.evaluate('L.ed.slide().notes'), notes + 'X', 'typing goes on where it was');
  u.clean();
}));

/* ---------- Phase 4: one owner per piece of state ---------- */

test('P3 Editing text, a click on the current slide\'s thumbnail ends the edit: Delete then deletes the slide', () => lectern(async (u) => {
  await setup(u, BOXES, 2);
  await editBeta(u);
  await u.click(await thumb(u, 0));
  assert.equal((await state(u)).editing, null, 'text editing ended');
  await u.key('Delete');
  assert.equal(await order(u), 'S2', 'the slide is deleted');
  u.clean();
}));

test('After an edit on the slide, the thumbnail\'s star previews the transition on the thumbnail shown', () => lectern(async (u) => {
  await setup(u, RECT);
  await u.evaluate(`(() => { L.ed.slide().trans = { type: 'fade', spd: 'slow', click: true }; L.panes.left.render();
    const play = L.show.previewTransition; L.show.previewTransition = function (host, ...a) { window.__previewOn = host.isConnected; return play(host, ...a); }; })()`);
  await pause(300);
  await u.click(await shapeBox(u, 'rectR'));
  await u.key('ArrowRight'); await pause(600);   // a nudge: the thumbnail is refreshed in place
  await u.click(await u.el('.th[data-i="0"] .th-star'));
  assert.equal(await u.evaluate('window.__previewOn'), true, 'the preview plays on the thumbnail on screen');
  u.clean();
}));

test('A selected diagram\'s frame: its edge shows the move cursor, a click there keeps it selected and a drag moves it', () => lectern(async (u) => {
  await setup(u, '');
  await u.evaluate(`L.diagram.insert(L.smartart.LAYOUTS[0].id)`); await pause(400);
  const id = await u.evaluate(`L.ed.slide().shapes[0].id`);
  assert.equal(await u.evaluate(`L.ed.slide().shapes[0].type`), 'group', 'setup: the diagram, selected');
  const geo0 = await geo(u, id);
  const fr = await u.evaluate(`(() => { const r = document.querySelector('.sel-frame[data-sid="${id}"]').getBoundingClientRect(); return { x: Math.round(r.left + r.width / 3), y: Math.round(r.top + 1), b: Math.round(r.bottom - 1) }; })()`);
  assert.equal(await u.evaluate(`getComputedStyle(document.elementFromPoint(${fr.x}, ${fr.y})).cursor`), 'move', 'the top edge');
  assert.equal(await u.evaluate(`getComputedStyle(document.elementFromPoint(${fr.x}, ${fr.b})).cursor`), 'move', 'the bottom edge');
  await u.click({ x: fr.x, y: fr.y });
  assert.deepEqual(await u.evaluate('L.ed.sel'), [id], 'a click on the edge keeps the diagram selected');
  await u.drag({ x: fr.x, y: fr.y }, { x: fr.x + 60, y: fr.y + 30 }, 6);
  const geo1 = await geo(u, id);
  assert.ok(geo1[0] > geo0[0] && geo1[1] > geo0[1], 'a drag on the edge moves it');
  assert.deepEqual(geo1.slice(2), geo0.slice(2), 'without resizing it');
  u.clean();
}));

test('P4a In the Slides pane, Delete twice deletes two slides', () => lectern(async (u) => {
  await setup(u, '', 4);
  await u.click(await thumb(u, 0));
  await u.key('Delete');
  assert.equal(await order(u), 'S2,S3,S4', 'setup: the first Delete');
  await u.key('Delete');
  assert.equal(await order(u), 'S3,S4', 'the second Delete deletes the next slide');
  u.clean();
}));

test('P4b After dragging a thumbnail to another place, Delete deletes the moved slide', () => lectern(async (u) => {
  await setup(u, '', 3);
  const a = await thumb(u, 0), c = await thumb(u, 2);
  await u.drag(a, { x: a.x, y: c.b + 10 }, 10);
  assert.equal(await order(u), 'S2,S3,S1', 'setup: the slide moved');
  await u.key('Delete');
  assert.equal(await order(u), 'S2,S3', 'the moved slide is deleted');
  u.clean();
}));

test('P6a Shift+click on a text box adds it to the selection, without editing it', () => lectern(async (u) => {
  await setup(u, BOXES);
  await u.click(await at(u, 'boxA', 3));
  await u.esc();
  assert.deepEqual(await state(u), { editing: null, sel: ['boxA'] }, 'setup: boxA selected as an object');
  await u.click(await at(u, 'boxB', 8), { shift: true });
  assert.deepEqual(await state(u), { editing: null, sel: ['boxA', 'boxB'] }, 'both boxes selected');
  u.clean();
}));

test('P6b Ctrl+click on a text box adds it to the selection, without editing it', () => lectern(async (u) => {
  await setup(u, `${BOXES}, ${RECT}`);
  await u.click(await shapeBox(u, 'rectR'));
  assert.deepEqual(await state(u), { editing: null, sel: ['rectR'] }, 'setup: the rectangle selected');
  await u.click(await at(u, 'boxB', 8), { ctrl: true });
  assert.deepEqual(await state(u), { editing: null, sel: ['rectR', 'boxB'] }, 'both selected');
  u.clean();
}));

test('P8 Dragging across the cells of a table selects cells; the table does not move', () => lectern(async (u) => {
  await setup(u, `table('tbl', 100, 150)`);
  await u.drag(await cell(u, 'tbl', 0, 0), await cell(u, 'tbl', 1, 2), 10);
  assert.deepEqual((await geo(u, 'tbl')).slice(0, 2), [100, 150], 'the table stays where it was');
  assert.deepEqual(await u.evaluate('L.ed.cellSel'), { r0: 0, c0: 0, r1: 1, c1: 2 }, 'the cells are selected (the Tables and Borders toolbar waited for the release, so the cells stayed under the pointer)');
  await pause(600);
  assert.equal(await u.evaluate('L.app.toolbarsEl.tables.hidden'), false, 'the Tables and Borders toolbar shows after the release');
  u.clean();
}));

test('P9 A click into one cell clears the highlight of a selected range of cells', () => lectern(async (u) => {
  await setup(u, `table('tbl', 100, 150)`);
  const highlights = () => u.evaluate(`document.querySelectorAll('.cell-sel').length`);
  await u.click(await cell(u, 'tbl', 0, 0));
  await u.drag(await cell(u, 'tbl', 0, 0), await cell(u, 'tbl', 1, 1), 10);
  assert.equal(await highlights(), 1, 'setup: the range is highlighted');
  await u.click(await cell(u, 'tbl', 2, 2));
  assert.equal(await highlights(), 0, 'no range highlighted');
  u.clean();
}));

test('P11 Editing text, a click on the hatched border selects the box as an object: Delete deletes it', () => lectern(async (u) => {
  await setup(u, BOXES);
  await u.click(await at(u, 'boxA', 3));
  const ring = await u.el('.edit-ring');
  await u.click({ x: ring.l + (ring.r - ring.l) * 0.3, y: ring.t + 1.5 });
  assert.deepEqual(await state(u), { editing: null, sel: ['boxA'] }, 'the box selected, its text not edited');
  await u.key('Delete');
  assert.equal(await text(u, 'boxA'), null, 'the box is deleted');
  u.clean();
}));

/** zoom to 200 % with the Zoom box, then drag the slide area's vertical scroll bar down */
async function zoomAndScroll(u) {
  await u.click(await u.el('#tb-zoom'));
  await u.type('200'); await u.key('Enter');
  /* the scroll bar where it is at the drag (a selected object's task pane may have opened beside it) */
  return async () => {
    const sb = await u.evaluate(`(() => { const s = L.ed.scroller, r = s.getBoundingClientRect(); return { x: r.right - (s.offsetWidth - s.clientWidth) / 2, y: r.top + 40 }; })()`);
    await u.press(sb); await u.glide(sb, { x: sb.x, y: sb.y + 120 }, 8); await u.release({ x: sb.x, y: sb.y + 120 });
    assert.ok(await u.evaluate('L.ed.scroller.scrollTop') > 0, 'setup: the slide area scrolled');
  };
}

test('P12a Dragging the scroll bar keeps the selection', () => lectern(async (u) => {
  await setup(u, `rect('rectR', 30, 30, 100, 60)`);
  const scroll = await zoomAndScroll(u);
  await u.click(await shapeBox(u, 'rectR'));
  assert.deepEqual((await state(u)).sel, ['rectR'], 'setup: the rectangle is selected');
  await scroll();
  assert.deepEqual(await state(u), { editing: null, sel: ['rectR'] }, 'still selected');
  u.clean();
}));

test('P12b Dragging the scroll bar keeps text editing', () => lectern(async (u) => {
  await setup(u, `box('boxA', 30, 30, 'Alpha beta gamma delta epsilon')`);
  const scroll = await zoomAndScroll(u);
  await u.click(await at(u, 'boxA', 3));
  assert.equal((await state(u)).editing, 'boxA', 'setup: editing');
  await scroll();
  assert.equal((await state(u)).editing, 'boxA', 'still editing');
  u.clean();
}));

test('P13 A double-click on a word in an unselected AutoShape selects the word', () => lectern(async (u) => {
  await setup(u, `rect('rectT', 80, 120, 420, 100, 'Alpha beta gamma delta')`);
  await u.dblclick(await at(u, 'rectT', 8));
  assert.equal((await state(u)).editing, 'rectT', 'its text is edited');
  assert.equal(await u.evaluate('String(window.getSelection())'), 'beta', 'the word is selected');
  u.clean();
}));

test('P4c A thumbnail drawn again after an edit keeps its elements, so a press on it in the meantime keeps its click', () => lectern(async (u) => {
  await setup(u, BOXES, 2);
  await u.evaluate(`window.__frame = document.querySelector('.th[data-i="0"] .th-frame')`);
  await u.click(await at(u, 'boxA', 3));
  await u.type('xyz');
  await u.esc();
  await pause(700);
  assert.equal(await u.evaluate('__frame.isConnected'), true, 'the thumbnail\'s frame is the same element');
  u.clean();
}));

test('A history step recorded in the middle of a drag cancels the drag first: the shape is back, one step to undo', () => lectern(async (u) => {
  await setup(u, RECT);
  await u.click(await shapeBox(u, 'rectR'));
  const b = await shapeBox(u, 'rectR');
  await u.press(b); await u.glide(b, { x: b.x + 80, y: b.y + 40 }, 6);
  await u.evaluate(`(() => { L.hist.push('Other'); L.ed.shape('rectR').fill = '#FF0000'; })()`);
  await u.release({ x: b.x + 80, y: b.y + 40 });
  assert.deepEqual(await geo(u, 'rectR'), [200, 150, 120, 80], 'the move was cancelled');
  assert.equal(await u.evaluate('L.ui.dragging()'), false, 'no drag is live');
  await u.key('z', { ctrl: true });
  assert.deepEqual(await geo(u, 'rectR'), [200, 150, 120, 80], 'Undo puts back the other edit, in the right place');
  assert.notEqual(await u.evaluate(`L.ed.shape('rectR').fill`), '#FF0000', 'the other edit is undone');
  u.clean();
}));

/* ---------- Phase 5: the rest ---------- */

test('P15 "Click to add first slide" in an empty presentation adds a slide', () => lectern(async (u) => {
  await u.evaluate(`(() => { L.pres.slides.length = 0; L.ed.goto(0); L.panes.left.render(); })()`);
  await pause(300);
  await u.click(await u.el('.ed-empty'));
  assert.equal(await u.evaluate('L.pres.slides.length'), 1, 'one slide');
  u.clean();
}));

test('A thumbnail dragged and let go over the slide moves no slide', () => lectern(async (u) => {
  await setup(u, '', 3);
  const a = await thumb(u, 0), mid = await slidePt(u, 480, 270);
  await u.drag(a, mid, 10);
  assert.equal(await order(u), 'S1,S2,S3', 'the order is kept');
  assert.equal(await u.evaluate(`document.querySelectorAll('.drop-line').length`), 0, 'no drop line left');
  u.clean();
}));

test('Typing over a range of table cells goes into the first cell and ends the range; Delete empties the range, one Undo brings it back', () => lectern(async (u) => {
  await setup(u, `table('tbl', 100, 150)`);
  const cellText = (r, c) => u.evaluate(`L.txt.paraText(L.ed.shape('tbl').tbl.rows[${r}].cells[${c}].tx.ps[0])`);
  await u.click(await cell(u, 'tbl', 0, 0));
  await u.drag(await cell(u, 'tbl', 0, 0), await cell(u, 'tbl', 1, 1), 10);
  assert.equal(await u.evaluate(`document.querySelectorAll('.cell-sel').length`), 1, 'setup: the range is highlighted');
  await u.type('Z');
  assert.equal(await u.evaluate(`document.querySelectorAll('.cell-sel').length`), 0, 'the highlight is gone');
  assert.equal(await u.evaluate('L.ed.cellSel'), null, 'the range ended');
  await u.key('Escape'); await pause(200);
  assert.match(await cellText(0, 0), /Z/, 'Z went into the first cell');
  assert.equal(await cellText(1, 1), 'r1c1', 'the other cells are as they were');

  await u.click(await cell(u, 'tbl', 1, 1));
  await u.drag(await cell(u, 'tbl', 1, 1), await cell(u, 'tbl', 2, 2), 10);
  await u.key('Delete');
  assert.deepEqual([await cellText(1, 1), await cellText(1, 2), await cellText(2, 1), await cellText(2, 2)], ['', '', '', ''], 'the range is empty');
  assert.equal(await cellText(1, 0), 'r1c0', 'a cell outside the range is kept');
  assert.equal(await u.evaluate('L.ed.cellSel'), null, 'the range ended');
  await u.key('z', { ctrl: true });
  assert.equal(await cellText(2, 2), 'r2c2', 'one Undo brings the cells back');
  u.clean();
}));

test('Slide show: the pen let go in another tab draws nothing after it', () => lectern(async (u) => {
  await setup(u, '', 2);
  await u.key('F5'); await pause(800);
  await u.key('p', { ctrl: true }); await pause(200);
  /** how many ink pixels lie on a horizontal line at client y between x0 and x1 */
  const inked = (y, x0, x1) => u.evaluate(`(() => { const c = document.querySelector('.show canvas.ink'), r = c.getBoundingClientRect(), k = c.width / r.width;
    const d = c.getContext('2d').getImageData(Math.round((${x0} - r.left) * k), Math.round((${y} - r.top) * k), Math.round((${x1} - ${x0}) * k), 1).data;
    let n = 0; for (let i = 3; i < d.length; i += 4) if (d[i]) n++; return n; })()`);
  await u.press({ x: 300, y: 300 }); await u.glide({ x: 300, y: 300 }, { x: 500, y: 300 }, 6);
  assert.ok(await inked(300, 320, 480) > 0, 'setup: the pen draws');
  await u.releaseInOtherTab();
  await u.move({ x: 300, y: 500 }); await u.glide({ x: 300, y: 500 }, { x: 600, y: 500 }, 8);
  assert.equal(await inked(500, 320, 580), 0, 'nothing drawn after the button was let go');
  assert.equal(await u.evaluate('L.show.active'), true, 'setup: still in the show');
  await u.esc(); await u.esc();
  u.clean();
}));
