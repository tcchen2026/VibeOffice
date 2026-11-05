// Ledger with real mouse and keyboard input (tests/lib/input.mjs). A test named after an audit item (S1…) checks what
// Excel does there and stays `todo` until the item is fixed (todos/UI_MOUSE_LOGIC_IMPROVE.md, Audit › Ledger); taking
// the mark off is part of the fix. Each test has its own tab, so a failing one leaves nothing behind.
//
//   node --test --test-concurrency=1 tests/ui/*.test.mjs
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { open, pause } from '../lib/input.mjs';

/* ---------------------------------------------------------------- helpers */
/** a new tab on Ledger */
const ledger = () => open('ledger');
/** [row, col] of an A1 reference */
const rc = (ref) => { const m = /^([A-Z]+)(\d+)$/.exec(ref); let c = 0; for (const ch of m[1]) c = c * 26 + ch.charCodeAt(0) - 64; return [+m[2] - 1, c - 1]; };
/** centre and box of a cell on screen (its sheet must be the one shown) */
const pt = (u, ref) => { const [r, c] = rc(ref); return u.evaluate(`(() => { const w = document.querySelector('#gridwrap').getBoundingClientRect(), b = L.grid.cellRect(${r}, ${c});
  return { x: Math.round(w.left + b.x + b.w / 2), y: Math.round(w.top + b.y + b.h / 2), l: w.left + b.x, t: w.top + b.y, r: w.left + b.x + b.w, b: w.top + b.y + b.h }; })()`); };
/** a column's header (letter) or a row's header (number) on screen */
const colHeader = async (u, ref) => { const p = await pt(u, ref); return { x: p.x, y: await u.evaluate(`Math.round(document.querySelector('#gridwrap').getBoundingClientRect().top + L.grid.frame().hdrH / 2)`) }; };
const rowHeader = async (u, ref) => { const p = await pt(u, ref); return { x: await u.evaluate(`Math.round(document.querySelector('#gridwrap').getBoundingClientRect().left + L.grid.frame().hdrW / 2)`), y: p.y }; };
/** what the formula bar shows for a cell (of sheet i, else the sheet shown) */
const text = (u, ref, i) => { const [r, c] = rc(ref); return u.evaluate(`(() => { const sh = ${i == null ? 'L.grid.sheet()' : `L.grid.wb.sheets[${i}]`}; return L.editor.cellText(sh, ${r}, ${c}); })()`); };
/** the selection: its ranges ('A1:B2 C3') and the active cell */
const selection = (u) => u.evaluate(`(() => { const s = L.grid.sel(), F = L.formula; return { ranges: s.ranges.map((x) => x.r1 === x.r2 && x.c1 === x.c2 ? F.cellName(x.r1, x.c1) : F.rangeName(x)).join(' '), active: F.cellName(s.r, s.c) }; })()`);
const sheetShown = (u) => u.evaluate(`document.querySelector('#tabs .stab.on').textContent`);
const tabNames = (u) => u.evaluate(`[...document.querySelectorAll('#tabs .stab')].map((t) => t.textContent)`);
const tab = (u, name) => u.evaluate(`(() => { const e = [...document.querySelectorAll('#tabs .stab')].find((x) => x.textContent === ${JSON.stringify(name)}); const r = e.getBoundingClientRect(); return { x: Math.round(r.left + r.width / 2), y: Math.round(r.top + r.height / 2) }; })()`);
/** the text of the message box or dialog on top, or null */
const dialog = (u) => u.evaluate(`(() => { const d = [...document.querySelectorAll('.dlg')].pop(); return d ? d.textContent.replace(/\\s+/g, ' ').trim() : null; })()`);
const menus = (u) => u.evaluate(`document.querySelectorAll('.menu').length`);
const nameBox = (u) => u.evaluate(`document.querySelector('#namebox').value`);
/** click a cell, type, Enter */
const enter = async (u, ref, s) => { await u.click(await pt(u, ref)); await u.type(s); await u.key('Enter'); };
/** the first drawn object on screen: its box, whether it shows as selected, its text */
const shape = (u) => u.evaluate(`(() => { const e = document.querySelector('#glayer .gobj'); if (!e) return null; const r = e.getBoundingClientRect(), t = e.querySelector('.gobj-t');
  return { x: Math.round(r.left + r.width / 2), y: Math.round(r.top + r.height / 2), l: Math.round(r.left), t: Math.round(r.top), selected: e.classList.contains('sel'), text: t ? t.textContent : '' }; })()`);
/** draw a rectangle from one cell to another with the Rectangle tool; it comes out selected */
const drawRect = async (u, from, to) => { await u.evaluate(`L.ui.exec('drawRect')`); await u.drag(await pt(u, from), await pt(u, to), 6); };
/** the fill handle of the selected cell */
const fillHandle = async (u, ref) => { const p = await pt(u, ref); return { x: Math.round(p.r - 1), y: Math.round(p.b - 1) }; };
/** a point on a container that is not on any control in it (the container itself is hit there) */
const blank = (u, selector) => u.evaluate(`(() => { const box = document.querySelector(${JSON.stringify(selector)}); const r = box.getBoundingClientRect();
  for (let x = Math.floor(r.right) - 2; x > r.left; x -= 1) for (const y of [r.top + r.height / 2, r.top + 3, r.bottom - 3]) {
    const e = document.elementFromPoint(x, y); if (e && (e === box || (box.contains(e) && !e.closest('button, input, select, textarea, a, .tb-btn, .tb-split, .combo, .stab, .mb-item')))) return { x, y: Math.round(y) }; }
  return null; })()`);
/** a mouse wheel turn over a point */
const wheel = async (u, p, dx, dy) => { await u.send('Input.dispatchMouseEvent', { type: 'mouseWheel', x: p.x, y: p.y, deltaX: dx, deltaY: dy }); await pause(300); };
/** the screen position of a cell, to see whether a pane scrolled */
const where = async (u, ref) => { const p = await pt(u, ref); return { x: p.x, y: p.y }; };

/* ---------------------------------------------------------------- S1: point mode in Edit mode */
test('S1a After F2, a click on a cell where the formula expects an operand inserts its reference', async () => {
  const u = await ledger();
  try {
    await enter(u, 'A1', '=1');
    await u.click(await pt(u, 'A1'));
    await u.key('F2');
    await u.type('+');
    await u.click(await pt(u, 'C3'));
    assert.equal(await dialog(u), null, 'no error message');
    await u.key('Enter');
    assert.equal(await text(u, 'A1'), '=1+C3');
    u.clean();
  } finally { await u.close(); }
});

test('S1b After a double-click, dragging over cells in a formula inserts the range', async () => {
  const u = await ledger();
  try {
    await u.dblclick(await pt(u, 'B5'));
    await u.type('=SUM(');
    await u.drag(await pt(u, 'D1'), await pt(u, 'D3'), 5);
    assert.equal(await dialog(u), null, 'no error message');
    await u.type(')');
    await u.key('Enter');
    assert.equal(await text(u, 'B5'), '=SUM(D1:D3)');
    u.clean();
  } finally { await u.close(); }
});

test('S1c Typing = in the formula bar and clicking a cell inserts its reference', async () => {
  const u = await ledger();
  try {
    await u.click(await pt(u, 'D6'));
    const bar = await u.el('#fbinput');
    await u.click({ x: bar.l + 20, y: bar.y });
    await u.type('=');
    await u.click(await pt(u, 'B2'));
    assert.equal(await dialog(u), null, 'no error message');
    await u.key('Enter');
    assert.equal(await text(u, 'D6'), '=B2');
    u.clean();
  } finally { await u.close(); }
});

/* ---------------------------------------------------------------- S2: Shift+Arrow grows from the active cell */
test('S2a Shift+Arrow extends from the active cell after the keyboard moved it', async () => {
  const u = await ledger();
  try {
    await u.click(await pt(u, 'A1'));
    await u.key('ArrowDown', { shift: true });
    await u.key('ArrowDown', { shift: true });
    await u.key('ArrowDown');
    assert.deepEqual(await selection(u), { ranges: 'A2', active: 'A2' }, 'Down leaves A2 selected');
    await u.key('ArrowRight', { shift: true });
    assert.deepEqual(await selection(u), { ranges: 'A2:B2', active: 'A2' });
    u.clean();
  } finally { await u.close(); }
});

test('S2b Shift+Arrow extends from a clicked cell, not from where an earlier drag ended', async () => {
  const u = await ledger();
  try {
    await u.drag(await pt(u, 'B2'), await pt(u, 'D5'), 6);
    await u.click(await pt(u, 'F10'));
    await u.key('ArrowDown', { shift: true });
    assert.deepEqual(await selection(u), { ranges: 'F10:F11', active: 'F10' });
    u.clean();
  } finally { await u.close(); }
});

test('S2c Shift+Right after a click on a column header adds the next column', async () => {
  const u = await ledger();
  try {
    await u.drag(await pt(u, 'B2'), await pt(u, 'D5'), 6);
    await u.click(await colHeader(u, 'H1'));
    await u.key('ArrowRight', { shift: true });
    assert.equal((await selection(u)).ranges, 'H:I');
    u.clean();
  } finally { await u.close(); }
});

/* ---------------------------------------------------------------- R4: a cell selected while a shape is */
for (const [id, how, go] of [
  ['R4a', 'Go To', async (u) => { await u.key('g', { ctrl: true }); await pause(300); await u.type('D5'); await u.key('Enter'); }],
  ['R4b', 'Find Next', async (u) => { await u.key('f', { ctrl: true }); await pause(300); await u.type('needle'); await u.key('Enter'); await pause(200); await u.esc(); }],
]) {
  test(`${id} With a shape selected, ${how} selects the cell, and Delete clears the cell, not the shape`, async () => {
    const u = await ledger();
    try {
      await enter(u, 'D5', 'needle');
      await drawRect(u, 'G3', 'I8');
      assert.equal((await shape(u)).selected, true, 'the new shape is selected');
      await go(u);
      await pause(300);
      assert.equal((await selection(u)).active, 'D5');
      assert.equal((await shape(u)).selected, false, 'the shape is no longer selected');
      await u.key('Delete');
      assert.equal(await text(u, 'D5'), '', 'Delete cleared D5');
      assert.equal(await u.evaluate('L.grid.sheet().drawings.length'), 1, 'the shape is still there');
      u.clean();
    } finally { await u.close(); }
  });
}

/* ---------------------------------------------------------------- S3: formula bar, then back into the cell */
test('S3 Typing in the cell after a visit to the formula bar is kept', async () => {
  const u = await ledger();
  try {
    await enter(u, 'A5', 'abc');
    await u.dblclick(await pt(u, 'A5'));
    const bar = await u.el('#fbinput');
    await u.click({ x: bar.r - 20, y: bar.y });
    await u.type('1');
    const box = await u.el('#celledit');
    await u.click({ x: box.l + 30, y: box.y });
    await u.key('End');
    await u.type('xyz');
    await u.key('Enter');
    assert.equal(await text(u, 'A5'), 'abc1xyz');
    u.clean();
  } finally { await u.close(); }
});

/* ---------------------------------------------------------------- S4: a shape's text, then a cell */
test('S4 After editing a shape\'s text, a click on a cell selects the cell and typing goes there', async () => {
  const u = await ledger();
  try {
    await drawRect(u, 'E5', 'G10');
    await u.dblclick(await shape(u));
    await u.type('Hi');
    await u.click(await pt(u, 'B3'));
    assert.equal((await shape(u)).selected, false, 'the shape is no longer selected');
    await u.type('42');
    await u.key('Enter');
    assert.equal(await text(u, 'B3'), '42', 'the typing went into B3');
    assert.equal((await shape(u)).text, 'Hi', 'the shape keeps its text');
    u.clean();
  } finally { await u.close(); }
});

/* ---------------------------------------------------------------- S5: the selection never leaves the cell being edited */
test('S5a A right-click on another cell while typing enters the typing first, as a click does', async () => {
  const u = await ledger();
  try {
    await u.click(await pt(u, 'A1'));
    await u.type('abc');
    await u.rightClick(await pt(u, 'C5'));
    assert.equal(await text(u, 'A1'), 'abc', 'A1 holds what was typed');
    assert.deepEqual(await selection(u), { ranges: 'C5', active: 'C5' });
    assert.equal(await menus(u), 1, 'the cell\'s shortcut menu is open');
    await u.esc();
    u.clean();
  } finally { await u.close(); }
});

test('S5b Going to a cell with the Name Box while typing enters the typing first', async () => {
  const u = await ledger();
  try {
    await u.click(await pt(u, 'A2'));
    await u.type('hello');
    await u.click(await u.el('#namebox'));
    await u.type('C5');
    await u.key('Enter');
    await u.type('X');
    await u.key('Enter');
    assert.deepEqual([await text(u, 'A2'), await text(u, 'C5')], ['hello', 'X'], 'A2 keeps "hello", the typing after the jump goes into C5');
    u.clean();
  } finally { await u.close(); }
});

/* ---------------------------------------------------------------- S6: a reference to another sheet */
test('S6 Enter after pointing at another sheet goes back to the formula\'s sheet', async () => {
  const u = await ledger();
  try {
    await u.click(await pt(u, 'A1'));
    await u.type('=');
    await u.click(await tab(u, 'Sheet2'));
    await u.click(await pt(u, 'B2'));
    await u.key('Enter');
    assert.equal(await sheetShown(u), 'Sheet1', 'back on Sheet1');
    assert.equal(await text(u, 'A1', 0), '=Sheet2!B2');
    assert.deepEqual(await selection(u), { ranges: 'A2', active: 'A2' }, 'Enter moved down from A1');
    u.clean();
  } finally { await u.close(); }
});

/* ---------------------------------------------------------------- S7: drag-and-drop onto data */
test('S7 Dropping a cell onto a cell with data asks before replacing it', async () => {
  const u = await ledger();
  try {
    await enter(u, 'A1', 'src');
    await enter(u, 'C1', 'keep me');
    await u.click(await pt(u, 'A1'));
    const a1 = await pt(u, 'A1'), c1 = await pt(u, 'C1');
    await u.drag({ x: Math.round(a1.r - 1), y: a1.y }, { x: Math.round(c1.r - 1), y: c1.y }, 8);   // by A1's right border
    assert.match(await dialog(u) || '', /replace/i, 'a message asks whether to replace the data');
    assert.equal(await text(u, 'C1'), 'keep me', 'nothing is replaced while it asks');
    await u.esc();
    assert.deepEqual([await text(u, 'A1'), await text(u, 'C1')], ['src', 'keep me'], 'Cancel keeps both');
    await u.drag({ x: Math.round(a1.r - 1), y: a1.y }, { x: Math.round(c1.r - 1), y: c1.y }, 8);
    await u.key('Enter');
    assert.deepEqual([await text(u, 'A1'), await text(u, 'C1')], ['', 'src'], 'OK moves the cell there');
    u.clean();
  } finally { await u.close(); }
});

test('Dropping a cell onto empty cells moves it without asking', async () => {
  const u = await ledger();
  try {
    await enter(u, 'A1', 'src');
    await u.click(await pt(u, 'A1'));
    const a1 = await pt(u, 'A1'), c1 = await pt(u, 'C1');
    await u.drag({ x: Math.round(a1.r - 1), y: a1.y }, { x: Math.round(c1.r - 1), y: c1.y }, 8);
    assert.equal(await dialog(u), null, 'no message');
    assert.deepEqual([await text(u, 'A1'), await text(u, 'C1')], ['', 'src']);
    u.clean();
  } finally { await u.close(); }
});

test('the Drawing toolbar\'s Line and Arrow draw in the drag\'s direction, the arrow with its head; AutoShapes offers every shape', async () => {
  const u = await ledger();
  try {
    await u.click(await u.button('drawLine'));
    await u.drag(await pt(u, 'C10'), await pt(u, 'F4'), 6);   // up and to the right
    await u.click(await u.button('drawArrow'));
    await u.drag(await pt(u, 'J4'), await pt(u, 'H9'), 6);    // down and to the left
    const ds = await u.evaluate(`JSON.stringify(L.grid.sheet().drawings.map((d) => [d.name, d.geom, !!d.flipH, !!d.flipV, d.tail ? d.tail.type : null]))`);
    assert.deepEqual(JSON.parse(ds), [['Line 2', 'line', false, true, null], ['Line 3', 'line', true, false, 'triangle']]);
    assert.equal(await u.evaluate(`document.querySelectorAll('#glayer marker').length`), 1, 'the arrowhead is drawn');
    await u.click(await u.el('[data-as]'));
    assert.deepEqual((await u.menuItems()).slice(0, 2), ['Lines', 'Basic Shapes']);
    await u.pick('Stars and Banners');
    const star = await u.evaluate(`(() => { const b = [...document.querySelectorAll('.menu .shape-grid button')].find((x) => x.getAttribute('aria-label') === '5-Point Star'); const r = b.getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height / 2 }; })()`);
    await u.click(star);
    await u.drag(await pt(u, 'C14'), await pt(u, 'E20'), 6);
    assert.equal(await u.evaluate(`L.grid.sheet().drawings[2].geom + ' / ' + L.grid.sheet().drawings[2].name`), 'star5 / 5-Point Star 4', 'named as Excel names it');
    u.clean();
  } finally { await u.close(); }
});

test('Esc while an AutoShape is dragged out draws nothing and ends the tool; the next drag selects cells', async () => {
  const u = await ledger();
  try {
    await u.evaluate(`L.ui.exec('drawRect')`);
    const a = await pt(u, 'C3'), b = await pt(u, 'E8');
    await u.press(a); await u.glide(a, b, 6); await u.esc(); await u.release(b);
    assert.equal(await u.evaluate('L.grid.sheet().drawings.length'), 0, 'no shape');
    await u.drag(await pt(u, 'B2'), await pt(u, 'C3'), 6);
    assert.equal((await selection(u)).ranges, 'B2:C3', 'the drag selects cells');
    assert.equal(await u.evaluate('L.grid.sheet().drawings.length'), 0, 'still no shape');
    u.clean();
  } finally { await u.close(); }
});

/* ---------------------------------------------------------------- S8: rename a sheet in its tab */
test('S8 A double-click on a sheet tab renames it in place', async () => {
  const u = await ledger();
  try {
    await u.dblclick(await tab(u, 'Sheet2'));
    assert.equal(await u.evaluate(`!!document.querySelector('#tabs input')`), true, 'the tab turns into a box with its name');
    await u.type('Data');
    await u.key('Enter');
    assert.deepEqual(await tabNames(u), ['Sheet1', 'Data', 'Sheet3']);
    u.clean();
  } finally { await u.close(); }
});

/* ---------------------------------------------------------------- S9: header drags that auto-scroll */
for (const [id, what, from, edge, axis] of [['S9a', 'column', 'C1', 'right', 'c'], ['S9b', 'row', 'A3', 'bottom', 'r']]) {
  test(`${id} Holding a ${what}-header drag past the ${edge} edge selects the ${what}s it scrolls to`, async () => {
    const u = await ledger();
    try {
      const geo = await u.evaluate(`(() => { const w = document.querySelector('#gridwrap').getBoundingClientRect(), f = L.grid.frame(), m = f.panes[f.panes.length - 1];
        return { right: Math.round(w.left + m.x + m.w - 1), bottom: Math.round(w.top + m.y + m.h - 1) }; })()`);
      const start = axis === 'c' ? await colHeader(u, from) : await rowHeader(u, from);
      const hold = axis === 'c' ? { x: geo.right, y: start.y } : { x: start.x, y: geo.bottom };
      const span = async () => { const s = await u.evaluate(`(() => { const s = L.grid.sel(), x = s.ranges[s.active]; return { end: x.${axis}2 }; })()`); return s.end; };
      const shown = (ref) => where(u, ref);   // a far cell's screen position shows how far the sheet scrolled
      await u.press(start);
      await u.glide(start, hold, 10);
      const end0 = await span(), pos0 = await shown('A1');
      await pause(1200);
      const end1 = await span(), pos1 = await shown('A1');
      await u.release(hold);
      const scrolled = axis === 'c' ? pos0.x - pos1.x : pos0.y - pos1.y;
      assert.ok(scrolled > 100, `the sheet scrolled (${scrolled} px)`);
      assert.ok(end1 - end0 >= 3, `the selection grew with it (last ${what} ${end0} → ${end1})`);
      u.clean();
    } finally { await u.close(); }
  });
}

/* ---------------------------------------------------------------- S10: Pick From List with the keyboard */
test('S10a Alt+Down, then the arrow keys and Enter, enter an entry from the list', async () => {
  const u = await ledger();
  try {
    for (const [ref, v] of [['A1', 'apple'], ['A2', 'banana'], ['A3', 'cherry']]) await enter(u, ref, v);
    await u.click(await pt(u, 'A4'));
    await u.key('ArrowDown', { alt: true });
    assert.equal(await u.evaluate(`!!document.querySelector('.acl')`), true, 'the list opens');
    await u.key('ArrowDown');
    await u.key('Enter');
    assert.ok(['apple', 'banana', 'cherry'].includes(await text(u, 'A4')), `A4 holds an entry from the list (it holds "${await text(u, 'A4')}")`);
    assert.equal((await selection(u)).active, 'A4', 'the active cell stays A4');
    assert.equal(await u.evaluate(`!!document.querySelector('.acl')`), false, 'the list is closed');
    u.clean();
  } finally { await u.close(); }
});

test('S10b Esc closes the Alt+Down list and leaves the cell as it was', async () => {
  const u = await ledger();
  try {
    for (const [ref, v] of [['A1', 'apple'], ['A2', 'banana']]) await enter(u, ref, v);
    await u.click(await pt(u, 'A3'));
    await u.key('ArrowDown', { alt: true });
    await u.esc();
    assert.equal(await u.evaluate(`!!document.querySelector('.acl')`), false, 'Esc closes the list');
    assert.deepEqual([await text(u, 'A3'), (await selection(u)).active], ['', 'A3']);
    u.clean();
  } finally { await u.close(); }
});

/* (Ctrl+PageDown and Ctrl+PageUp never reach a page in a Chrome tab, which switches tabs with them) */
test('A sheet tab\'s context menu closed with Esc: the arrow keys still move on the sheet', async () => {
  const u = await ledger();
  try {
    await u.click(await pt(u, 'B2'));
    await u.rightClick(await tab(u, 'Sheet1'));
    assert.ok(await menus(u), 'setup: the tab\'s menu is open');
    await u.esc(); await pause(150);
    await u.key('ArrowDown');
    assert.equal((await selection(u)).active, 'B3');
    u.clean();
  } finally { await u.close(); }
});

/* ---------------------------------------------------------------- S11: chrome keeps the keyboard on the sheet */
for (const [id, what, where0] of [
  ['S11a', 'the status bar', (u) => blank(u, '#statusbar')],
  ['S11b', 'the empty part of the sheet-tab strip', (u) => blank(u, '#tabs')],
  ['S11c', 'a toolbar\'s background', (u) => blank(u, '#tb-formatting')],
  ['S11d', 'a sheet-tab scroll arrow', (u) => u.el('#tabnav button')],
]) {
  test(`${id} After a click on ${what}, the arrow keys and Delete still work on the sheet`, async () => {
    const u = await ledger();
    try {
      await enter(u, 'C4', 'data');
      await u.click(await pt(u, 'C3'));
      const p = await where0(u);
      assert.ok(p, 'found a point on ' + what);
      await u.click(p);
      await u.key('ArrowDown');
      assert.equal((await selection(u)).active, 'C4', 'Down moves to C4');
      await u.key('Delete');
      assert.equal(await text(u, 'C4'), '', 'Delete clears C4');
      u.clean();
    } finally { await u.close(); }
  });
}

/* ---------------------------------------------------------------- S12: a release the page never saw */
test('S12a A cell drag whose release was lost does not follow the pointer any more', async () => {
  const u = await ledger();
  try {
    const b3 = await pt(u, 'B3'), d6 = await pt(u, 'D6'), f12 = await pt(u, 'F12');
    await u.press(b3);
    await u.glide(b3, d6, 5);
    await u.loseRelease(d6);
    const after = await selection(u);
    await u.glide(d6, f12, 6);   // no button held
    assert.deepEqual(await selection(u), after, 'moving without a button changes nothing');
    u.clean();
  } finally { await u.close(); }
});

test('S12b A fill-handle drag whose release was lost leaves nothing behind', async () => {
  const u = await ledger();
  try {
    await enter(u, 'A1', '1');
    await u.click(await pt(u, 'A1'));
    const h = await fillHandle(u, 'A1'), a5 = await pt(u, 'A5'), a12 = await pt(u, 'A12');
    await u.press(h);
    await u.glide(h, a5, 5);
    await u.loseRelease(a5);
    await u.glide(a5, a12, 5);
    await u.click(await pt(u, 'E10'));
    assert.equal(await nameBox(u), 'E10', 'the Name Box shows the clicked cell, not a fill size');
    assert.deepEqual(await Promise.all(['A2', 'A5', 'A12'].map((r) => text(u, r))), ['', '', ''], 'nothing was filled');
    u.clean();
  } finally { await u.close(); }
});

test('S12c A shape move whose release was lost is cancelled, not dropped by the next click', async () => {
  const u = await ledger();
  try {
    await drawRect(u, 'E5', 'G10');
    const s0 = await shape(u);
    await u.press(s0);
    await u.glide(s0, { x: s0.x + 40, y: s0.y + 20 }, 4);
    await u.loseRelease({ x: s0.x + 40, y: s0.y + 20 });
    const far = await pt(u, 'M26');
    await u.glide({ x: s0.x + 40, y: s0.y + 20 }, far, 6);
    await u.click(far);
    const s1 = await shape(u);
    assert.deepEqual([s1.l, s1.t], [s0.l, s0.t], 'the shape stays where it was');
    u.clean();
  } finally { await u.close(); }
});

test('An edit recorded in the middle of a shape move ends the move first: the move records nothing', async () => {
  const u = await ledger();
  try {
    await drawRect(u, 'E5', 'G10');
    const s0 = await shape(u), to = { x: s0.x + 60, y: s0.y + 40 };
    await u.press(s0);
    await u.glide(s0, to, 4);
    await u.evaluate(`L.ops.tx(L.grid.wb, 'Typing', () => L.ops.enter(L.grid.sheet(), 0, 0, 'x'))`);   // an edit from outside the drag
    await u.glide(to, { x: to.x + 20, y: to.y }, 2);
    await u.release({ x: to.x + 20, y: to.y });
    const s1 = await shape(u);
    assert.deepEqual([s1.l, s1.t], [s0.l, s0.t], 'the shape stays where it was');
    await u.key('z', { ctrl: true });
    assert.equal(await text(u, 'A1'), '', 'one Undo takes back the edit');
    const s2 = await shape(u);
    assert.deepEqual([s2.l, s2.t], [s0.l, s0.t], 'and the shape is still where it was');
    u.clean();
  } finally { await u.close(); }
});

/* ---------------------------------------------------------------- S13: frozen panes */
test('S13a A drag within the frozen rows does not scroll the rows below them', async () => {
  const u = await ledger();
  try {
    await u.click(await pt(u, 'C3'));
    await u.evaluate(`L.ui.exec('freezePanes')`);
    const top = await where(u, 'A60');
    await wheel(u, await pt(u, 'F20'), 0, 33 * 30);   // 30 rows down
    const before = await where(u, 'A60');
    assert.ok(top.y - before.y > 300, 'the wheel scrolled the rows below the frozen ones');
    const h1 = await pt(u, 'H1'), j2 = await pt(u, 'J2');
    await u.press(h1);
    await u.glide(h1, j2, 6);
    await pause(800);
    await u.release(j2);
    assert.deepEqual(await selection(u), { ranges: 'H1:J2', active: 'H1' });
    assert.deepEqual(await where(u, 'A60'), before, 'the rows below the frozen ones stayed where they were');
    u.clean();
  } finally { await u.close(); }
});

test('S13b A drag within the frozen columns does not scroll the columns beside them', async () => {
  const u = await ledger();
  try {
    await u.click(await pt(u, 'C3'));
    await u.evaluate(`L.ui.exec('freezePanes')`);
    const left = await where(u, 'Z5');
    await u.send('Input.dispatchMouseEvent', { type: 'mouseWheel', x: 600, y: 400, deltaX: 0, deltaY: 33 * 6, modifiers: 8 });   // Shift+wheel: 6 columns right
    await pause(300);
    const before = await where(u, 'Z5');
    assert.ok(left.x - before.x > 200, 'Shift+wheel scrolled the columns beside the frozen ones');
    const a5 = await pt(u, 'A5'), a9 = await pt(u, 'A9');
    await u.press(a5);
    await u.glide(a5, a9, 6);
    await pause(800);
    await u.release(a9);
    assert.deepEqual(await selection(u), { ranges: 'A5:A9', active: 'A5' });
    assert.deepEqual(await where(u, 'Z5'), before, 'the columns beside the frozen ones stayed where they were');
    u.clean();
  } finally { await u.close(); }
});

/* ---------------------------------------------------------------- S14 */
test('S14a Holding the button down on a hyperlink selects the cell without following the link', async () => {
  const u = await ledger();
  try {
    // a link from A1 to Sheet3!B2 (Insert ▸ Hyperlink, Place in This Document)
    await u.evaluate(`(() => { const sh = L.grid.sheet(); L.ops.enter(sh, 0, 0, 'Go to Sheet3'); sh.links.push({ ref: { r1: 0, c1: 0, r2: 0, c2: 0 }, location: 'Sheet3!B2' }); L.grid.paint(true); })()`);
    await u.click(await pt(u, 'E5'));
    const a1 = await pt(u, 'A1');
    await u.press(a1);
    await pause(1500);
    await u.release(a1);
    assert.equal(await sheetShown(u), 'Sheet1', 'still on Sheet1');
    assert.deepEqual(await selection(u), { ranges: 'A1', active: 'A1' }, 'A1 is selected');
    u.clean();
  } finally { await u.close(); }
});

test('S14b Shift+click while pointing in a formula extends the reference to a range', async () => {
  const u = await ledger();
  try {
    await u.click(await pt(u, 'A10'));
    await u.type('=');
    await u.click(await pt(u, 'B2'));
    await u.click(await pt(u, 'D5'), { shift: true });
    await u.key('Enter');
    assert.equal(await text(u, 'A10'), '=B2:D5');
    u.clean();
  } finally { await u.close(); }
});

test('S14c Esc during a fill-handle drag cancels the fill', async () => {
  const u = await ledger();
  try {
    await enter(u, 'A4', '1');
    await u.click(await pt(u, 'A4'));
    const h = await fillHandle(u, 'A4'), a8 = await pt(u, 'A8');
    await u.press(h);
    await u.glide(h, a8, 6);
    await u.esc();
    await u.release(a8);
    assert.deepEqual(await Promise.all(['A5', 'A6', 'A7', 'A8'].map((r) => text(u, r))), ['', '', '', '']);
    u.clean();
  } finally { await u.close(); }
});

test('S14d The right button pressed during a left-button drag opens no shortcut menu', async () => {
  const u = await ledger();
  try {
    const b10 = await pt(u, 'B10'), d14 = await pt(u, 'D14'), e16 = await pt(u, 'E16');
    await u.press(b10);
    await u.glide(b10, d14, 5);
    await u.mouse('mousePressed', d14, { button: 'right' }); await pause(80);
    await u.mouse('mouseReleased', d14, { button: 'right' }); await pause(200);
    assert.equal(await menus(u), 0, 'no shortcut menu while dragging');
    await u.glide(d14, e16, 5);
    await u.release(e16);
    assert.equal(await menus(u), 0, 'none after the drag either');
    u.clean();
  } finally { await u.close(); }
});

test('A drop-down inside a dialog opens above it: Format Cells ▸ Patterns ▸ Pattern picks a pattern', async () => {
  const u = await ledger();
  const centre = (sel) => u.evaluate(`(() => { const e = [...document.querySelectorAll(${JSON.stringify(sel)})].find((x) => x.offsetParent); if (!e) return null; const r = e.getBoundingClientRect();
    return { x: Math.round(r.left + r.width / 2), y: Math.round(r.top + r.height / 2), top: document.elementFromPoint(r.left + r.width / 2, r.top + r.height / 2) === e }; })()`);
  try {
    await u.click(await pt(u, 'B2'));
    await u.evaluate(`L.dlg.formatCells(4)`); await pause(300);
    await u.click(await u.evaluate(`(() => { const b = [...document.querySelectorAll('.dlg button')].find((x) => x.textContent.trim() === 'Pattern'); const r = b.getBoundingClientRect(); return { x: Math.round(r.left + r.width / 2), y: Math.round(r.top + r.height / 2) }; })()`));
    const swatch = await centre('.menu .pat-grid button[data-pattern="darkGrid"]');
    assert.ok(swatch, 'the pattern menu is open');
    assert.ok(swatch.top, 'its swatches are on top, not under the dialog');
    await u.click(swatch);
    await u.click(await u.evaluate(`(() => { const b = [...document.querySelectorAll('.dlg button')].find((x) => x.textContent.trim() === 'OK'); const r = b.getBoundingClientRect(); return { x: Math.round(r.left + r.width / 2), y: Math.round(r.top + r.height / 2) }; })()`));
    assert.equal(await dialog(u), null, 'the dialog is closed');
    assert.equal(await u.evaluate(`(() => { const s = L.grid.sheet(); return (L.grid.wb.styles.get(s.styleAt(1, 1)).fill || {}).pattern; })()`), 'darkGrid');
    u.clean();
  } finally { await u.close(); }
});

test('Save As ▸ Tools opens its menu under the button, above the dialog', async () => {
  const u = await ledger();
  try {
    await u.evaluate(`void L.dlg.saveAs()`); await pause(300);
    const b = await u.evaluate(`(() => { const e = [...document.querySelectorAll('.dlg button')].find((x) => x.textContent.trim() === 'Tools'); const r = e.getBoundingClientRect(); return { x: Math.round(r.left + r.width / 2), y: Math.round(r.top + r.height / 2), l: r.left, b: r.bottom }; })()`);
    await u.click(b);
    const m = await u.evaluate(`(() => { const e = document.querySelector('.menu'); if (!e) return null; const r = e.getBoundingClientRect(), i = e.querySelector('.m-item').getBoundingClientRect();
      return { l: r.left, t: r.top, top: document.elementFromPoint(i.left + i.width / 2, i.top + i.height / 2).closest('.menu') === e, text: e.textContent }; })()`);
    assert.ok(m, 'the menu is open');
    assert.ok(Math.abs(m.l - b.l) < 3 && Math.abs(m.t - b.b) < 3, `under the button (menu at ${m.l},${m.t}; button ${b.l},${b.b})`);
    assert.ok(m.top, 'above the dialog');
    assert.match(m.text, /General Options/);
    u.clean();
  } finally { await u.close(); }
});
