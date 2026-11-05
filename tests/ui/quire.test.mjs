// Quire with real mouse and keyboard input (tests/lib/input.mjs). A test named after an audit item (Q1…) checks what
// Word does there and stays `todo` until the item is fixed (todos/UI_MOUSE_LOGIC_IMPROVE.md); the rest must pass.
//
//   node --test --test-concurrency=1 tests/ui/*.test.mjs
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { open, pause } from '../lib/input.mjs';

/* ---------- the audit (todos/UI_MOUSE_LOGIC_IMPROVE.md, Audit › Quire): one test per item, each in a tab of its own,
   so a failing one leaves no dialog, menu or focus behind ---------- */

/** fn(u) in a new Quire tab on a document made by `build` (page code; D, and d the new document, in scope; it may
    await), the caret at the start; then no browser errors */
async function inQuire(build, fn, opts) {
  const u = await open('quire', opts);
  try {
    await u.evaluate(`(async () => {
      const { D } = L, d = D.newDoc();
      ${build}
      L.app.loadDoc(d, 'UI test');
      await new Promise((r) => setTimeout(r, 600));
      const p = D.firstPara(d.main.blocks); L.ed.setSel(D.pos(p, 0)); L.ed.focus();
    })()`);
    await fn(u);
    u.clean();
  } finally { await u.close(); }
}
/* documents (page code for inQuire) */
const PARAS = (list) => `d.main.blocks = ${JSON.stringify(list)}.map((s) => D.para([D.text(s)]));`;
/** a floating rectangle in front of the text, on page 1, anchored at the start of block b */
const SHAPE = (b) => `{ const it = L.drawing.newShape('rect', 90, 60); L.drawing.setWrap(it, 'none'); it.name = 'Rectangle 1';
  it.float.posH = { rel: 'page', off: 330 }; it.float.posV = { rel: 'page', off: 200 }; d.main.blocks[${b}].runs.unshift(it); }`;
/** an inline picture (100 × 60 pt) at offset o of block b */
const PICTURE = (b, o) => `{ const c = document.createElement('canvas'); c.width = 133; c.height = 80; const g = c.getContext('2d'); g.fillStyle = '#3a7'; g.fillRect(0, 0, 133, 80);
  const id = L.media.add(await new Promise((r) => c.toBlob(r, 'image/png')), 'pic.png'); const p = d.main.blocks[${b}], s = p.runs[0].text;
  p.runs = [D.text(s.slice(0, ${o})), D.item('img', { media: id, w: 100, h: 60, name: 'pic.png', alt: '' }), D.text(s.slice(${o}))]; }`;
/** a 2 × 2 table, "cell 00 text" …, inserted as block b */
const TABLE = (b) => `{ const tb = D.simpleTable(2, 2, 432); tb.rows.forEach((r, i) => r.cells.forEach((c, j) => { c.blocks = [D.para([D.text('cell ' + i + j + ' text')])]; }));
  d.main.blocks.splice(${b}, 0, tb); }`;
const HEADER = `d.hf.h1 = { kind: 'hdr', id: 'h1', blocks: [D.para([D.text('Header text')], { style: 'Header' })] }; d.sect.refs = { hdr: { default: 'h1' }, ftr: {} };`;
/* what the user sees */
/** screen point just inside the character at offset o of paragraph p (page code; B is the body's blocks) */
const pointIn = (u, p, o) => u.evaluate(`(() => { const B = L.D.doc.main.blocks, m = L.layout.posToDom(L.D.pos(${p}, ${o})), r = document.createRange();
  r.setStart(m.node, m.offset); r.collapse(true); const box = r.getClientRects()[0] || r.getBoundingClientRect(); return { x: box.left + 1, y: box.top + box.height / 2 }; })()`);
const atB = (u, b, o) => pointIn(u, `B[${b}]`, o);
/** a point in the left margin (the selection bar) beside the line at offset o of block b */
const margin = async (u, b, o) => { const p = await atB(u, b, o); return { x: p.x - 30, y: p.y }; };
/** the body: paragraph texts, a table as rows of cell texts */
const body = (u) => u.evaluate(`L.D.doc.main.blocks.map((b) => b.t === 'tbl' ? b.rows.map((r) => r.cells.map((c) => c.blocks.map((x) => L.D.ptext(x)).join('/'))) : L.D.ptext(b))`);
/** the text the selection covers (paragraphs joined by \n), what typing replaces */
const selected = (u) => u.evaluate(`(() => { const s = L.ed.sel; if (!s) return ''; const [a, b] = L.D.order(L.D.doc, s.a, s.f);
  return L.D.parasBetween(L.D.doc, a.p, b.p).map((p) => L.D.ptext(p, p === a.p ? a.o : 0, p === b.p ? b.o : L.D.plen(p))).join('\\n'); })()`);
/** the body block the selection ends in */
const selBlock = (u) => u.evaluate('L.D.doc.main.blocks.indexOf(L.ed.sel.f.p)');
/** the text highlighted in the page */
const highlighted = (u) => u.evaluate('String(getSelection())');
const undoSteps = (u) => u.evaluate('L.D.histState().undo.length');
const undoLabels = (u) => u.evaluate('L.D.histState().undo.map((x) => x.label)');
const shapes = (u) => u.evaluate(`document.querySelectorAll('#pages .flt').length`);
const editingHeader = (u) => u.evaluate(`L.layout.root.classList.contains('hfmode')`);
const near = (a, b) => Math.abs(a.l - b.l) < 1 && Math.abs(a.t - b.t) < 1 && Math.abs(a.r - b.r) < 1 && Math.abs(a.b - b.b) < 1;
/* ---------- what works today (these must pass) ---------- */

/** [text, has prop] for each run of body block b */
const runs = (u, b, prop) => u.evaluate(`L.D.doc.main.blocks[${b}].runs.map((r) => [r.text, !!(r.rPr && r.rPr.${prop})])`);
const objectState = (u) => u.evaluate(`({ handles: document.querySelectorAll('.obj-handles').length, outlined: document.querySelectorAll('.sel-obj').length, selected: !!L.ed.objSel })`);
const TEXT = ['Source plain target words follow here', 'Second paragraph of ordinary text for clicking', 'Third paragraph with more words to paint'];

test('the highlighter marks the whole drag, as one undo step', () => inQuire(PARAS(TEXT), async (u) => {
  await u.click(await u.button('highlightBtn'));
  assert.equal(await u.painting(), true, 'the Highlight button with no selection arms the highlighter');
  const undo0 = await undoSteps(u);
  await u.drag(await atB(u, 0, 7), await atB(u, 0, 25));
  assert.deepEqual(await runs(u, 0, 'hl'), [['Source ', false], ['plain target words', true], [' follow here', false]]);
  assert.equal(await undoSteps(u) - undo0, 1, 'one undo step');
  await u.esc();
  assert.equal(await u.painting(), false, 'Esc puts the highlighter down');
}, { look: 'classic' }));

for (const leave of ['a click into text', 'Esc']) {
  test(`a drawn shape's handles go after ${leave}, and a click selects it again`, () => inQuire(PARAS(TEXT), async (u) => {
    await u.click(await u.button('drawRect'));
    const corner = await atB(u, 1, 10);
    await u.drag({ x: corner.x + 300, y: corner.y + 40 }, { x: corner.x + 400, y: corner.y + 120 }, 4);
    assert.deepEqual(await objectState(u), { handles: 1, outlined: 1, selected: true }, 'the drawn shape is selected');
    if (leave === 'Esc') await u.esc(); else await u.click(await atB(u, 1, 5));
    assert.deepEqual(await objectState(u), { handles: 0, outlined: 0, selected: false }, 'no handles');
    await u.click({ x: corner.x + 350, y: corner.y + 80 });
    assert.deepEqual(await objectState(u), { handles: 1, outlined: 1, selected: true }, 'selected again');
  }, { look: 'classic' }));
}

/* ---------- the audit items ---------- */

const FIND = ['Alpha beta gamma.', 'Second beta paragraph.', 'Third beta paragraph.'];
test('Q1 Find Next keeps the keyboard in the Find box: a second Enter finds the next match and never edits the text', () => inQuire(PARAS(FIND), async (u) => {
  await u.key('f', { ctrl: true }); await pause(300);
  await u.type('beta');
  await u.key('Enter'); await u.key('Enter');
  assert.deepEqual(await body(u), FIND, 'the document is unchanged');
  assert.equal(await u.focused(), 'fr-find', 'the Find box has the keyboard');
  assert.deepEqual([await selBlock(u), await selected(u)], [1, 'beta'], 'the second "beta" is selected');
  await u.esc();
  await u.type('X');
  assert.equal((await body(u))[1], 'Second X paragraph.', 'after the box closes, typing replaces the found text');
}));

test('Find and Replace: Replace goes one match at a time and keeps the keyboard; typing after Esc replaces the next match', () => inQuire(PARAS(FIND), async (u) => {
  await u.key('h', { ctrl: true }); await pause(300);
  await u.type('beta');
  await u.click(await u.el('#fr-repl')); await u.type('X');
  const replace = await u.evaluate(`(() => { const b = [...document.querySelectorAll('.dlg button.btn')].find((b) => b.textContent === 'Replace'); const r = b.getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height / 2 }; })()`);
  for (let i = 0; i < 3; i++) await u.click(replace);
  assert.deepEqual(await body(u), ['Alpha X gamma.', 'Second X paragraph.', 'Third beta paragraph.'], 'the first click finds, then each replaces one');
  assert.notEqual(await u.focused(), 'pages', 'the dialog keeps the keyboard');
  await u.esc();
  await u.type('Y');
  assert.equal((await body(u))[2], 'Third Y paragraph.', 'typing replaces the match found next');
}));

test('R2 Find after a floating shape was selected: the found word is the selection, typing replaces it, the shape stays', () => inQuire(PARAS(FIND) + SHAPE(2), async (u) => {
  const text = await body(u);
  await u.click(await u.el('#pages .flt'));
  const shape0 = await u.el('#pages .flt');
  await u.key('f', { ctrl: true }); await pause(300);
  await u.type('beta'); await u.key('Enter');     // the next "beta" after the shape's anchor, in the third paragraph
  await u.esc();
  await u.type('X');
  await u.key('ArrowRight');
  assert.equal((await body(u))[2], text[2].replace('beta', 'X'), 'typing replaced the found word');
  assert.equal(await shapes(u), 1, 'the shape is still there');
  assert.ok(near(await u.el('#pages .flt'), shape0), 'and where it was');
}));

test('R3 after a press let go in another window, moving the caret into a template placeholder still selects it', () => inQuire(PLACEHOLDER, async (u) => {
  await u.press(await atB(u, 0, 2));
  await u.releaseInOtherTab();
  let s = '';
  for (let i = 0; i < 8 && !s; i++) { await u.key('ArrowRight'); s = await selected(u); }
  assert.equal(s, '[Click here and type name]', 'the whole placeholder is selected');
}));

test('R3b a text selection let go in another window paints nothing, and the painter stays on', () => inQuire(PARAS(TEXT), async (u) => {
  await u.click(await atB(u, 0, 2));
  await u.click(await u.button('painter'));
  await u.press(await atB(u, 1, 7)); await u.glide(await atB(u, 1, 7), await atB(u, 1, 16), 5);
  await u.releaseInOtherTab();
  await u.move(await atB(u, 2, 3));
  assert.deepEqual(await runs(u, 1, 'b'), [['Second paragraph of ordinary text for clicking', false]], 'nothing painted');
  assert.equal(await u.painting(), true, 'the painter is still on');
  await u.esc();
}, { look: 'classic' }));

const SHAPE_DOC = ['First paragraph of ordinary text for clicking.', 'Second paragraph of ordinary text.', 'Third paragraph.'];
test('Q2 Enter or Tab on a selected floating shape keeps the shape and leaves the text alone', () => inQuire(PARAS(SHAPE_DOC) + SHAPE(1), async (u) => {
  const text = await body(u);   // the shape's anchor is a character of block 1
  for (const k of ['Enter', 'Tab']) {
    await u.click(await atB(u, 0, 5));
    await u.click(await u.el('#pages .flt'));
    await u.key(k);
    assert.equal(await shapes(u), 1, `the shape is still there after ${k}`);
    assert.deepEqual(await body(u), text, `${k} puts no paragraph or tab into the text`);
  }
}));

const PLACEHOLDER = `d.main.blocks = [D.para([D.text('Dear '), D.item('fb', { fid: 'f1', instr: ' MACROBUTTON  DoFieldClick [Click here and type name] ' }), D.item('fs', { fid: 'f1' }),
  D.text('[Click here and type name]'), D.item('fe', { fid: 'f1' }), D.text(', thank you for the long letter you sent.')]), D.para([D.text('Second paragraph.')])];`;
test('Q3 a drag that starts inside a template placeholder selects from there, and never moves the placeholder', () => inQuire(PLACEHOLDER, async (u) => {
  const before = await body(u), undo0 = await undoSteps(u);
  await u.click(await atB(u, 1, 3));
  await u.drag(await atB(u, 0, 15), await atB(u, 0, 45));   // from "[Click he|re" to "thank you fo|r"
  assert.deepEqual(await body(u), before, 'the text is where it was');
  assert.equal(await undoSteps(u), undo0, 'nothing to undo');
  assert.match(await selected(u), /, thank you/, 'the selection runs from the placeholder to where the mouse was let go');
}));

test('a click on a template placeholder selects all of it, and typing replaces it', () => inQuire(PLACEHOLDER, async (u) => {
  await u.click(await atB(u, 0, 12));
  assert.equal(await selected(u), '[Click here and type name]', 'the whole placeholder is selected');
  await u.type('Ann');
  assert.match((await body(u))[0], /^Dear Ann, thank you/);
}));

const DND = ['Alpha beta gamma delta.', 'Second paragraph of ordinary text.', 'Third paragraph.'];
test('Q4 after a text drag dropped outside the document, text dragged in from the font box is inserted and nothing else moves', () => inQuire(PARAS(DND), async (u) => {
  await u.drag(await atB(u, 0, 6), await atB(u, 0, 10), 4);         // select "beta"
  await u.drag(await atB(u, 0, 7), await u.el('#statusbar'), 10);   // drop it on the status bar: nothing happens
  assert.equal((await body(u))[0], DND[0], 'a drop outside the document changes nothing');
  const box = await u.el('#tb-font');
  await u.click(box);                                               // the font's name is selected in its box
  const name = await u.evaluate(`document.querySelector('#tb-font').value`);
  await u.drag(box, await atB(u, 2, 6), 12);                        // and dragged into "Third |paragraph."
  const text = await body(u);
  assert.equal(text[0], DND[0], '"beta" stays where it is');
  assert.ok(name && text[2].includes(name), `the name "${name}" is in the third paragraph: "${text[2]}"`);
}));

test('Q5 a word moved forward in its paragraph: typing next goes at the moved word', () => inQuire(PARAS(['Alpha beta gamma delta epsilon zeta eta theta.', 'Second paragraph.']), async (u) => {
  await u.drag(await atB(u, 0, 6), await atB(u, 0, 11), 4);    // "beta "
  await u.drag(await atB(u, 0, 7), await atB(u, 0, 23), 10);   // dropped before "epsilon"
  assert.equal((await body(u))[0], 'Alpha gamma delta beta epsilon zeta eta theta.', 'moved');
  await u.type('X');
  /* Word selects the moved text, so X replaces it; X just after it would do too */
  assert.match((await body(u))[0], /^Alpha gamma delta (beta )?Xepsilon zeta eta theta\.$/);
}));
/** the text of block b as it shows with Track Changes: deleted runs left out */
const shown = (u, b) => u.evaluate(`L.D.doc.main.blocks[${b}].runs.filter((r) => !(r.rPr && r.rPr.del)).map((r) => r.text || '').join('')`);
test('Q5 tracked: a word moved forward in its paragraph lands at the drop caret (the deleted source stays in place)', () => inQuire(PARAS(['Alpha beta gamma delta epsilon zeta eta theta.', 'Second paragraph.']) + 'd.settings.track = true;', async (u) => {
  await u.drag(await atB(u, 0, 6), await atB(u, 0, 11), 4);    // "beta "
  await u.drag(await atB(u, 0, 7), await atB(u, 0, 23), 10);   // dropped before "epsilon"
  assert.equal(await shown(u, 0), 'Alpha gamma delta beta epsilon zeta eta theta.');
}));

const FOCUS = ['First paragraph of text.', 'Second paragraph of ordinary text.', 'Third paragraph.'];
const FOCUS_Z = ['First paragraph of text.', 'Second Z of ordinary text.', 'Third paragraph.'];
for (const [id, close, how] of [['Q6a', 'Escape', 'Esc'], ['Q6b', 'Enter', 'Enter (OK)']]) {
  test(`${id} the Font dialog (Ctrl+D) closed with ${how}: the selection is kept and typing replaces it`, () => inQuire(PARAS(FOCUS), async (u) => {
    await u.drag(await atB(u, 1, 7), await atB(u, 1, 16), 6);   // "paragraph"
    await u.key('d', { ctrl: true }); await pause(300);
    assert.ok(await u.el('.dlg'), 'the Font dialog is open');
    await u.key(close); await pause(200);
    await u.type('Z');
    assert.deepEqual(await body(u), FOCUS_Z);
  }));
}
test('Q6c View ▸ Ruler chosen with the mouse: typing still goes to the document, replacing the selection', () => inQuire(PARAS(FOCUS), async (u) => {
  await u.drag(await atB(u, 1, 7), await atB(u, 1, 16), 6);
  try {
    await u.menu('View', 'Ruler');
    await u.type('Z');
    assert.deepEqual(await body(u), FOCUS_Z);
  } finally {   // the ruler setting is kept for the profile: put it back
    await u.evaluate(`if (!L.app.opts.ruler) { L.app.opts.ruler = true; L.app.saveOpts(); L.app.updateRulers(); }`);
  }
}, { look: 'classic' }));
test('Q6d Format ▸ Paragraph… from the menu bar, closed with Esc: typing replaces the selection', () => inQuire(PARAS(FOCUS), async (u) => {
  await u.drag(await atB(u, 1, 7), await atB(u, 1, 16), 6);
  await u.menu('Format', 'Paragraph');
  assert.ok(await u.el('.dlg'), 'the Paragraph dialog is open');
  await u.esc(); await pause(200);
  await u.type('Z');
  assert.deepEqual(await body(u), FOCUS_Z);
}, { look: 'classic' }));

const SHORT = Array.from({ length: 12 }, (_, i) => `Paragraph ${i} with some words.`);
const LONG = Array.from({ length: 70 }, (_, i) => `Paragraph ${i} with some words.`);
/** page i's header (client px) */
const hdrBox = (u, i) => u.evaluate(`(() => { const r = document.querySelectorAll('#pages .pg')[${i}].querySelector('.pg-hdr').getBoundingClientRect(); return { l: r.left, r: r.right, t: r.top, b: r.bottom, y: r.top + r.height / 2 }; })()`);
/** scroll page i to the top of the window, as the scroll bar would */
const toPage = (u, i) => u.evaluate(`(async () => { document.querySelectorAll('#pages .pg')[${i}].scrollIntoView({ block: 'start' }); await new Promise((r) => setTimeout(r, 300)); })()`);
/** the body block the text of page i starts with */
const firstOnPage = (u, i) => u.evaluate(`L.D.doc.main.blocks.indexOf(L.D.byId(L.D.doc, +document.querySelectorAll('#pages .pg')[${i}].querySelector('.pg-body .p').dataset.pid))`);

test('Q7 while the header is edited, a click on the dimmed body text leaves the body alone (Enter, Delete, Backspace, letters)', () => inQuire(PARAS(SHORT) + HEADER, async (u) => {
  const h = await hdrBox(u, 0);
  await u.dblclick({ x: h.r - 40, y: h.y });
  assert.ok(await editingHeader(u), 'editing the header');
  await u.click(await atB(u, 3, 4));
  for (const k of ['Enter', 'Delete', 'Backspace']) await u.key(k);
  await u.type('Q');
  assert.deepEqual(await body(u), SHORT, 'the body is not edited');
}));

test('Q8a a double-click on page 2\'s header opens it there: the view stays on page 2, the caret where double-clicked', () => inQuire(PARAS(LONG) + HEADER, async (u) => {
  await toPage(u, 1);
  const h = await hdrBox(u, 1);
  await u.dblclick({ x: h.r - 40, y: h.y });   // to the right of "Header text"
  assert.ok(await editingHeader(u), 'editing the header');
  const sc = await u.el('#scroller'), h2 = await hdrBox(u, 1);
  assert.ok(h2.t >= sc.t && h2.b <= sc.b, `page 2's header is still on screen (at ${Math.round(h2.t)}, the window ${Math.round(sc.t)}–${Math.round(sc.b)})`);
  await u.type('H');
  assert.equal(await u.evaluate(`L.D.ptext(L.D.doc.hf.h1.blocks[0])`), 'Header textH', 'typing goes where the double-click was');
}));
test('Q8b Esc closes the header and the caret is back where it was in the body', () => inQuire(PARAS(LONG) + HEADER, async (u) => {
  await toPage(u, 1);
  const n = (await firstOnPage(u, 1)) + 3;
  await u.click(await atB(u, n, 4));
  const h = await hdrBox(u, 1);
  await u.dblclick({ x: h.r - 40, y: h.y });
  assert.ok(await editingHeader(u), 'editing the header');
  await u.esc();
  assert.equal(await editingHeader(u), false, 'Esc closes the header');
  await u.type('Z');
  assert.equal((await body(u))[n], `ParaZgraph ${n} with some words.`);
}));
test('Q8c a double-click on the body closes the header, with the caret where double-clicked', () => inQuire(PARAS(SHORT) + HEADER, async (u) => {
  const h = await hdrBox(u, 0);
  await u.dblclick({ x: h.r - 40, y: h.y });
  assert.ok(await editingHeader(u), 'editing the header');
  const line = await atB(u, 3, 0), pageBody = await u.el('#pages .pg .pg-body');
  await u.dblclick({ x: pageBody.r - 30, y: line.y });   // to the right of "Paragraph 3 with some words."
  assert.equal(await editingHeader(u), false, 'the header is closed');
  await u.type('Z');
  assert.equal((await body(u))[3], 'Paragraph 3 with some words.Z');
}));
test('Q8d a double-click in the top margin opens the (empty) header for editing', () => inQuire(PARAS(SHORT), async (u) => {
  const m = await u.evaluate(`(() => { const p = document.querySelector('#pages .pg'), pr = p.getBoundingClientRect(), b = p.querySelector('.pg-body').getBoundingClientRect(); return { x: b.left + 100, y: (pr.top + b.top) / 2 }; })()`);
  await u.dblclick(m);
  assert.ok(await editingHeader(u), 'editing the header');
  await u.type('H');
  assert.match(await u.evaluate(`document.querySelector('#pages .pg .pg-hdr').textContent`), /H/, 'the typed text is in the header');
  assert.deepEqual(await body(u), SHORT, 'the body is not edited');
}));

test('Q9 a right-click in a table cell, the caret outside the table: the first menu already has the table commands', () => inQuire(PARAS(['First paragraph of text.', 'Second paragraph.']) + TABLE(1), async (u) => {
  await u.click(await atB(u, 0, 3));
  await u.rightClick(await pointIn(u, 'B[1].rows[0].cells[0].blocks[0]', 3));
  assert.ok((await u.menuItems()).includes('Table Properties...'), `the menu: ${(await u.menuItems()).join(', ')}`);
  await u.esc();
}));

const PIC = ['First paragraph of text.', 'Picture here and text after it.', 'Third paragraph with words.'];
test('Q10 a right-click on a picture, selected or not, selects it (handles shown) and offers Format Picture…', () => inQuire(PARAS(PIC) + PICTURE(1, 5), async (u) => {
  const handles = () => u.evaluate(`document.querySelectorAll('.obj-handles').length`);
  const pic = await u.el('#pages .img');
  await u.click(pic);
  assert.equal(await handles(), 1, 'a click selects the picture');
  await u.rightClick(pic);
  assert.equal(await handles(), 1, 'a selected picture keeps its handles');
  assert.ok((await u.menuItems()).includes('Format Picture...'), 'the menu has Format Picture...');
  await u.esc();
  await u.click(await atB(u, 0, 3));
  await u.rightClick(pic);
  assert.equal(await handles(), 1, 'an unselected picture is selected by the right-click');
  assert.ok((await u.menuItems()).includes('Format Picture...'), 'the menu has Format Picture...');
  await u.esc();
}));

test('Q11 Ctrl+click selects the sentence clicked, so typing replaces it', () => inQuire(PARAS(['First sentence here. Second sentence follows. Third one ends it.', 'Another paragraph.']), async (u) => {
  await u.click(await atB(u, 0, 3));
  await u.click(await atB(u, 0, 50), { ctrl: true });   // in "Third one ends it."
  await u.type('X');
  assert.equal((await body(u))[0], 'First sentence here. Second sentence follows. X');
}));

const ALPHA = 'Alpha beta gamma delta epsilon zeta eta theta iota kappa lambda mu nu xi omicron pi rho sigma tau upsilon phi chi psi omega.';
const SELBAR = [`${ALPHA} ${ALPHA} ${ALPHA}`, 'Second paragraph here. It has two sentences.', 'Third paragraph text.'];
test('Q12a a drag down the selection bar selects whole lines, down to the line at the pointer', () => inQuire(PARAS(SELBAR), async (u) => {
  await u.drag(await margin(u, 0, 0), await margin(u, 1, 0), 10);
  const s = await selected(u);
  assert.ok(s.startsWith('Alpha beta') && s.includes('It has two sentences.'), `selected: "${s.slice(0, 30)}…${s.slice(-30)}"`);
}));
test('Q12b Shift+click in the selection bar extends the selection to that line', () => inQuire(PARAS(SELBAR), async (u) => {
  await u.click(await atB(u, 2, 5));
  await u.click(await margin(u, 0, 0), { shift: true });
  const s = await selected(u);
  assert.ok(s.startsWith('Alpha beta') && s.includes('Second paragraph here') && s.endsWith('Third'), `selected: "${s.slice(0, 30)}…${s.slice(-30)}"`);
}));
test('Q12c a triple-click in the selection bar selects the whole document', () => inQuire(PARAS(SELBAR), async (u) => {
  await u.clicks(await margin(u, 1, 0), 3);
  const s = await selected(u);
  assert.ok(s.startsWith('Alpha beta') && s.includes('Third paragraph text.'), `selected: "${s.slice(0, 30)}…${s.slice(-30)}"`);
}));

test('Q13 a click that doesn\'t move on an indent marker, a tab stop or the rotate handle changes nothing',
  () => inQuire(PARAS(SHAPE_DOC) + `d.main.blocks[0].pPr = { ind: { l: 25, fl: 0 }, tabs: [{ pos: 100, al: 'left' }] };` + SHAPE(1), async (u) => {
    await u.click(await atB(u, 0, 5));
    const pPr = () => u.evaluate(`JSON.stringify(L.D.doc.main.blocks[0].pPr)`);
    const before = await pPr(), undo0 = await undoSteps(u);
    await u.click(await u.el('.rmark.ind-left'));
    await u.click(await u.el('.rmark.tabstop'));
    assert.equal(await pPr(), before, 'the indent and the tab stop are where they were');
    assert.equal(await undoSteps(u), undo0, 'nothing to undo after the ruler clicks');
    await u.click(await u.el('#pages .flt'));
    await u.click(await u.el('.obj-handles .oh-rot'));
    assert.equal(await undoSteps(u), undo0, 'nothing to undo after the rotate handle click');
  }));

test('Q14a Esc while a shape is being drawn (button still down) cancels it: nothing is inserted', () => inQuire(PARAS(SHAPE_DOC), async (u) => {
  await u.click(await u.button('drawRect'));
  const a = await atB(u, 1, 10), from = { x: a.x + 300, y: a.y + 40 }, to = { x: a.x + 400, y: a.y + 120 };
  await u.press(from); await u.glide(from, to, 5);
  await u.esc();
  await u.release(to);
  assert.equal(await shapes(u), 0, 'no shape');
  assert.equal(await u.evaluate(`document.querySelectorAll('.draw-ghost').length`), 0, 'no outline left');
  assert.equal(await u.evaluate(`L.$('#scroller').classList.contains('drawing-mode')`), false, 'Esc also ends the tool');
}, { look: 'classic' }));
test('Q14b Esc while a shape is being moved cancels the move', () => inQuire(PARAS(SHAPE_DOC) + SHAPE(1), async (u) => {
  const b0 = await u.el('#pages .flt > *'), to = { x: b0.x + 80, y: b0.y + 60 }, undo0 = await undoSteps(u);
  await u.press(b0); await u.glide(b0, to, 6);
  await u.esc();
  await u.release(to);
  assert.ok(near(await u.el('#pages .flt > *'), b0), 'the shape is where it was');
  assert.equal(await undoSteps(u), undo0, 'nothing to undo');
}));
test('Q14c Esc while a picture is being resized cancels the resize', () => inQuire(PARAS(PIC) + PICTURE(1, 5), async (u) => {
  await u.click(await u.el('#pages .img'));
  const p0 = await u.el('#pages .img'), h = await u.el('.obj-handles .oh-se'), to = { x: h.x + 40, y: h.y + 30 };
  await u.press(h); await u.glide(h, to, 6);
  await u.esc();
  await u.release(to);
  const p1 = await u.el('#pages .img');
  assert.deepEqual([Math.round(p1.r - p1.l), Math.round(p1.b - p1.t)], [Math.round(p0.r - p0.l), Math.round(p0.b - p0.t)], 'the picture keeps its size');
}));

test('Q15 a click on a floating shape ends the text selection: no text stays highlighted', () => inQuire(PARAS(SHAPE_DOC) + SHAPE(1), async (u) => {
  await u.drag(await atB(u, 0, 6), await atB(u, 0, 15), 6);
  assert.equal(await highlighted(u), 'paragraph', 'the word is selected');
  await u.click(await u.el('#pages .flt'));
  assert.ok(!(await highlighted(u)).includes('paragraph'), 'the word is no longer highlighted');
}));

test('Q16 a drag from a table cell down into the text selects whole rows, and Delete removes them', () => inQuire(PARAS(['First paragraph of text.', 'Second paragraph.', 'Third paragraph after the table.']) + TABLE(1), async (u) => {
  await u.drag(await pointIn(u, 'B[1].rows[0].cells[0].blocks[0]', 2), await atB(u, 3, 5), 8);
  assert.equal(await u.evaluate(`document.querySelectorAll('#pages td.csel').length`), 4, 'every cell is highlighted');
  await u.key('Delete');
  const text = await body(u);
  assert.ok(!text.some(Array.isArray), 'the table is gone');
  assert.equal(text[0], 'First paragraph of text.');
}));
test('Q16 tracked: Delete over the rows and the text below marks them deleted, and Reject All brings everything back', () => inQuire(PARAS(['First paragraph of text.', 'Second paragraph.', 'Third paragraph after the table.']) + TABLE(1) + 'd.settings.track = true;', async (u) => {
  const before = await body(u);
  await u.drag(await pointIn(u, 'B[1].rows[0].cells[0].blocks[0]', 2), await atB(u, 3, 5), 8);
  await u.key('Delete');
  await u.evaluate(`L.ui.exec('rejectAll')`); await pause(300);
  assert.deepEqual(await body(u), before);
}));

/** insert the Diagram Gallery's pyramid (Level 1, 2, 3) at the caret; its shapes on screen */
const insertDiagram = async (u) => {
  await u.evaluate(`L.ui.exec('insertDiagram')`); await pause(400);
  await u.click(await u.evaluate(`(() => { const b = document.querySelector('.dlg [aria-label="Pyramid Diagram"]'); const r = b.getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height / 2 }; })()`));
  await u.click(await u.el('.dlg .dlg-foot .btn.primary')); await pause(700);
  return u.evaluate(`[...document.querySelectorAll('#pages .shp.sgrp .shp')].map((e) => { const r = e.getBoundingClientRect(); return { x: Math.round(r.left + r.width / 2), y: Math.round(r.top + r.height / 2), l: Math.round(r.left), t: Math.round(r.top), text: e.textContent.trim() }; })`);
};
test('A diagram\'s labels take a caret: a click on "Level 2" and typing edit that label, and the edit is saved', () => inQuire(PARAS(['First paragraph.', 'Second paragraph.']), async (u) => {
  await u.click(await atB(u, 0, 5));
  const kids = await insertDiagram(u);
  assert.deepEqual(kids.map((k) => k.text), ['Level 1', 'Level 2', 'Level 3'], 'setup: the pyramid');
  const label = await u.evaluate(`(() => { const t = [...document.querySelectorAll('#pages .shp.sgrp .tbx .pc')].find((e) => e.textContent.includes('Level 2')); const r = t.getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height / 2 }; })()`);
  await u.click(label); await pause(200);
  await u.key('End'); await u.type('Z');
  const labels = () => u.evaluate(`[...document.querySelectorAll('#pages .shp.sgrp .shp')].map((e) => e.textContent.trim())`);
  assert.deepEqual(await labels(), ['Level 1', 'Level 2Z', 'Level 3'], 'typing goes into the label');
  assert.equal((await body(u))[0], 'First\ufffc paragraph.', 'and not into the paragraph the diagram is in');
  const saved = await u.evaluate(`(async () => { const blob = await L.docx.write(L.D.doc); const d = await L.docx.read(await blob.arrayBuffer()); const g = (d.doc || d).main.blocks[0].runs.find((r) => r.t === 'group'); return L.D.groupTextboxes(g).map((tb) => tb.blocks.map((p) => L.D.ptext(p)).join('/')); })()`);
  assert.deepEqual(saved, ['Level 1', 'Level 2Z', 'Level 3'], 'saved and read back');
}));

test('The move cursor over a picture, shape or diagram (a drag moves it), the text cursor over its text', () => inQuire(PARAS(['First paragraph.', 'Second paragraph.']), async (u) => {
  await u.click(await atB(u, 0, 5));
  const kids = await insertDiagram(u);
  const cursorAt = (p) => u.evaluate(`getComputedStyle(document.elementFromPoint(${p.x}, ${p.y})).cursor`);
  assert.equal(await cursorAt({ x: kids[2].l + 30, y: kids[2].t + 4 }), 'move', 'on a shape of the diagram');
  const label = await u.evaluate(`(() => { const t = [...document.querySelectorAll('#pages .shp.sgrp .tbx .pc')].find((e) => e.textContent.includes('Level 3')); const r = t.getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height / 2 }; })()`);
  assert.equal(await cursorAt(label), 'text', 'on its text');
}));

test('Q17a with "Drag-and-drop text editing" off, a drag from selected text selects instead of moving', () => inQuire(PARAS(DND) + 'L.app.opts.dragDrop = false;', async (u) => {
  await u.drag(await atB(u, 0, 6), await atB(u, 0, 10), 4);    // "beta"
  await u.drag(await atB(u, 0, 7), await atB(u, 2, 6), 10);    // from inside it to "Third |paragraph."
  assert.deepEqual(await body(u), DND, 'nothing moves');
  assert.ok((await selected(u)).endsWith('Third '), 'a new selection, to where the mouse was let go');
}));
test('Q17b an inline picture dragged to another place in the text moves there', () => inQuire(PARAS(PIC) + PICTURE(1, 5), async (u) => {
  const pic = await u.el('#pages .img');
  await u.click(pic);
  await u.drag(pic, await atB(u, 2, 6), 10);
  assert.equal(await u.evaluate(`L.D.doc.main.blocks.findIndex((b) => b.runs.some((r) => r.t === 'img'))`), 2, 'the picture is in the third paragraph');
}));
test('Q17c a right-click during a shape move cancels the move, and no menu opens', () => inQuire(PARAS(SHAPE_DOC) + SHAPE(1), async (u) => {
  const b0 = await u.el('#pages .flt > *'), to = { x: b0.x + 80, y: b0.y + 60 }, undo0 = await undoSteps(u);
  await u.press(b0); await u.glide(b0, to, 6);
  await u.press(to, { button: 'right' }); await u.release(to, { button: 'right' });
  await u.release(to);
  assert.equal(await u.evaluate(`document.querySelectorAll('.menu').length`), 0, 'no menu');
  assert.ok(near(await u.el('#pages .flt > *'), b0), 'the shape is where it was');
  assert.equal(await undoSteps(u), undo0, 'nothing to undo');
}));

/* the structure under the fixes (Phases 4 and 5): one way into and out of an object selection, the text press on
   ui.drag, and no edit in the middle of a drag */
test('any text selection ends an object selection: Ctrl+Home after a floating shape is clicked, then typing goes to the text', () => inQuire(PARAS(SHAPE_DOC) + SHAPE(1), async (u) => {
  await u.click(await u.el('#pages .flt'));
  await u.key('Home', { ctrl: true });
  await u.type('Z');
  assert.equal((await body(u))[0], 'Z' + SHAPE_DOC[0], 'typed at the start of the document');
  assert.deepEqual(await objectState(u), { handles: 0, outlined: 0, selected: false }, 'the shape is no longer selected');
}));
test('a right-click in the middle of a text selection ends it: the painter paints nothing and stays on, no menu opens', () => inQuire(PARAS(TEXT), async (u) => {
  await u.click(await atB(u, 0, 2));
  await u.click(await u.button('painter'));
  const a = await atB(u, 1, 7), b = await atB(u, 1, 16);
  await u.press(a); await u.glide(a, b, 5);
  await u.mouse('mousePressed', b, { button: 'right' }); await pause(60); await u.mouse('mouseReleased', b, { button: 'right' }); await pause(200);
  await u.release(b);
  assert.deepEqual(await runs(u, 1, 'b'), [['Second paragraph of ordinary text for clicking', false]], 'nothing painted');
  assert.equal(await u.evaluate(`document.querySelectorAll('.menu').length`), 0, 'no menu');
  assert.equal(await u.painting(), true, 'the painter is still on');
  await u.esc();
}, { look: 'classic' }));
test('an edit in the middle of a shape move ends the move first: the shape stays, the edit is the only step', () => inQuire(PARAS(SHAPE_DOC) + SHAPE(1), async (u) => {
  const b0 = await u.el('#pages .flt > *'), to = { x: b0.x + 80, y: b0.y + 60 }, undo0 = await undoSteps(u);
  await u.press(b0); await u.glide(b0, to, 6);
  await u.evaluate(`L.ed.edit('Typing', () => L.O.insertText(L.D.pos(L.D.doc.main.blocks[2], 0), 'Q', {}))`);   // as a timer or another window would
  await u.release(to);
  assert.ok(near(await u.el('#pages .flt > *'), b0), 'the shape is where it was');
  assert.deepEqual((await undoLabels(u)).slice(undo0), ['Typing'], 'only the edit is recorded');
  assert.equal((await body(u))[2], 'Q' + SHAPE_DOC[2]);
}));

/* what Quire's drags still do on the shared drag lifecycle (ui.drag): each commits on release, as one undo step */
test('drags on a shape move, resize and rotate it; a picture\'s handle resizes it', () => inQuire(PARAS(PIC) + PICTURE(1, 5) + SHAPE(2), async (u) => {
  const box = () => u.el('#pages .flt > *');
  const b0 = await box();
  await u.drag(b0, { x: b0.x + 60, y: b0.y + 40 });
  const b1 = await box();
  assert.ok(Math.abs(b1.l - b0.l - 60) < 2 && Math.abs(b1.t - b0.t - 40) < 2, `moved with the pointer (${Math.round(b1.l - b0.l)}, ${Math.round(b1.t - b0.t)})`);
  await u.click(b1);
  let h = await u.el('.obj-handles .oh-se');
  await u.drag(h, { x: h.x + 30, y: h.y + 20 });
  const b2 = await box();
  assert.ok(b2.r - b2.l > b1.r - b1.l + 20 && b2.b - b2.t > b1.b - b1.t + 10, 'bigger');
  await u.click(b2);
  h = await u.el('.obj-handles .oh-rot');
  await u.drag(h, { x: b2.r + 60, y: b2.y }, 10);
  assert.match(await u.evaluate(`document.querySelector('#pages .flt > *').style.transform`), /rotate\(90deg\)/, 'turned a quarter to the right');
  await u.click(await u.el('#pages .img'));
  const p0 = await u.el('#pages .img');
  h = await u.el('.obj-handles .oh-se');
  await u.drag(h, { x: h.x + 40, y: h.y + 30 });
  const p1 = await u.el('#pages .img');
  assert.ok(p1.r - p1.l > p0.r - p0.l + 30, 'the picture is wider');
  assert.deepEqual(await undoLabels(u), ['Move Object', 'Resize Object', 'Rotate', 'Resize Picture']);
}));

test('the drawing tool: a drag draws the shape that size, a click puts one of the default size', () => inQuire(PARAS(SHAPE_DOC), async (u) => {
  const a = await atB(u, 1, 10);
  await u.click(await u.button('drawRect'));
  await u.drag({ x: a.x + 300, y: a.y + 40 }, { x: a.x + 400, y: a.y + 120 }, 5);
  const s = await u.el('#pages .flt > *');
  assert.ok(Math.abs(s.r - s.l - 100) < 2 && Math.abs(s.b - s.t - 80) < 2, `the size dragged out (${Math.round(s.r - s.l)} × ${Math.round(s.b - s.t)})`);
  await u.click(await atB(u, 0, 3));
  await u.click(await u.button('drawRect'));
  await u.click({ x: a.x + 50, y: a.y + 160 });
  assert.equal(await shapes(u), 2, 'a click draws one too');
  assert.deepEqual(await undoLabels(u), ['Insert AutoShape', 'Insert AutoShape']);
}, { look: 'classic' }));

test('the ruler: a drag moves the left indent or a tab stop, a click sets a tab stop, a tab stop dragged off goes; Esc puts a marker back', () => inQuire(PARAS(SHAPE_DOC) + `d.main.blocks[0].pPr = { ind: { l: 18 }, tabs: [{ pos: 108, al: 'left' }] };`, async (u) => {
  await u.click(await atB(u, 0, 5));
  const pPr = () => u.evaluate(`L.D.doc.main.blocks[0].pPr`);
  let m = await u.el('.rmark.ind-left');
  await u.drag(m, { x: m.x + 30, y: m.y });
  const l1 = (await pPr()).ind.l;
  assert.ok(l1 > 36, `the left indent follows the drag (${l1} pt)`);
  m = await u.el('.rmark.ind-left');
  await u.press(m); await u.glide(m, { x: m.x + 30, y: m.y }); await u.esc(); await u.release({ x: m.x + 30, y: m.y });
  assert.equal((await pPr()).ind.l, l1, 'Esc: the indent stays');
  assert.ok(Math.abs((await u.el('.rmark.ind-left')).x - m.x) < 1, 'and its marker goes back');
  assert.equal(await u.evaluate(`document.querySelectorAll('.r-guide').length`), 0, 'no guide line left');
  m = await u.el('.rmark.tabstop');
  await u.drag(m, { x: m.x + 40, y: m.y });
  const tab = (await pPr()).tabs.map((t) => t.pos);
  assert.ok(tab.length === 1 && tab[0] > 130, `the tab stop follows the drag (${tab})`);
  m = await u.el('.rmark.tabstop');
  await u.drag(m, { x: m.x, y: m.y + 60 });
  assert.deepEqual((await pPr()).tabs, [], 'dragged off the ruler, it goes');
  const r = await u.el('#hruler');
  await u.click({ x: m.x + 60, y: r.y });
  assert.equal((await pPr()).tabs.length, 1, 'a click on the ruler sets a tab stop');
  assert.deepEqual(await undoLabels(u), ['Left Indent', 'Move Tab', 'Clear Tab', 'Set Tab']);
}));

test('a table\'s column border: a drag sets the widths; a click or Esc leaves them', () => inQuire(PARAS(['First paragraph of text.', 'Second paragraph.']) + TABLE(1), async (u) => {
  const grid = () => u.evaluate(`L.D.doc.main.blocks[1].grid`);
  const g0 = await grid();
  let td = await u.el('#pages td[data-cid]');
  await u.click({ x: td.r, y: td.y });
  assert.deepEqual(await grid(), g0, 'a click changes nothing');
  await u.press({ x: td.r, y: td.y }); await u.glide({ x: td.r, y: td.y }, { x: td.r + 40, y: td.y }); await u.esc(); await u.release({ x: td.r + 40, y: td.y });
  assert.deepEqual(await grid(), g0, 'Esc: no change');
  assert.equal(await u.evaluate(`document.querySelectorAll('.r-guide').length`), 0, 'no guide line left');
  td = await u.el('#pages td[data-cid]');
  await u.drag({ x: td.r, y: td.y }, { x: td.r + 40, y: td.y });
  const g1 = await grid();
  assert.ok(g1[0] > g0[0] + 20 && Math.abs(g1[0] + g1[1] - g0[0] - g0[1]) < 0.1, `the first column is wider, the table as wide (${g1})`);
  assert.deepEqual(await undoLabels(u), ['Column Width']);
}));

test('a table\'s move handle: over the table; a click selects it, a drag floats it where it is dropped, Esc leaves it, a right-click opens its menu', () => inQuire(PARAS(['First paragraph of text.', 'Second paragraph.', 'Third paragraph, long enough to go on beside a narrow table that was moved over here by its handle.']) + TABLE(1) + 'd.main.blocks[1].grid = [90, 90];', async (u) => {
  const tbl = () => u.el('#pages table.tbl');
  const float = () => u.evaluate(`JSON.stringify(L.D.doc.main.blocks.find((b) => b.t === 'tbl').tblPr.float || null)`).then(JSON.parse);
  let t0 = await tbl();
  await u.move({ x: t0.l + 20, y: t0.t + 8 });
  let h = await u.el('.tbl-move');
  assert.ok(h && h.r <= t0.l + 1 && h.b <= t0.t + 1, 'the handle shows outside its top left corner');
  await u.click({ x: h.x, y: h.y });
  assert.equal(await u.evaluate(`L.ed.cellSel ? L.ed.cellSel.cells.length : 0`), 4, 'a click selects every cell');
  assert.equal(await undoSteps(u), 0);
  await u.press({ x: h.x, y: h.y }); await u.glide({ x: h.x, y: h.y }, { x: h.x + 120, y: h.y + 90 }); await u.esc(); await u.release({ x: h.x + 120, y: h.y + 90 });
  assert.equal(await float(), null, 'Esc: still in the text');
  assert.equal(await u.evaluate(`document.querySelectorAll('.tbl-ghost').length`), 0, 'no outline left');
  assert.equal(await undoSteps(u), 0);
  await u.move({ x: t0.l + 20, y: t0.t + 8 });
  h = await u.el('.tbl-move');
  await u.drag({ x: h.x, y: h.y }, { x: h.x + 150, y: h.y + 60 });
  const f = await float();
  assert.ok(f && f.vAnchor === 'text' && f.hAnchor === 'margin' && f.l === 9 && f.r === 9, `floating as Word makes it (${JSON.stringify(f)})`);
  const t1 = await tbl();
  assert.ok(Math.abs(t1.l - t0.l - 150) <= 2 && Math.abs(t1.t - t0.t - 60) <= 2, `where it was dropped (${t1.l - t0.l}, ${t1.t - t0.t})`);
  assert.deepEqual(await undoLabels(u), ['Move Table']);
  const xml = await u.evaluate(`(async () => { const pk = await L.opc.open(new Uint8Array(await (await L.docx.write(L.D.doc)).arrayBuffer())); return (pk.text('word/document.xml').match(/<w:tblpPr[^>]*>/) || [''])[0]; })()`);
  assert.match(xml, /^<w:tblpPr w:leftFromText="180" w:rightFromText="180" w:vertAnchor="text" w:horzAnchor="margin" w:tblpX="\d+" w:tblpY="-?\d+"\/>$/, 'saved as Word writes a dragged table');
  await u.move({ x: t1.l + 20, y: t1.t + 8 });
  h = await u.el('.tbl-move');
  await u.rightClick({ x: h.x, y: h.y });
  assert.ok((await u.menuItems()).some((x) => /Table Properties/.test(x)), 'the table\'s menu');
  await u.esc();
  await u.key('z', { ctrl: true });
  assert.equal(await float(), null, 'Undo: back in the text');
  t0 = await tbl();
  assert.ok(Math.abs(t0.l - (t1.l - 150)) <= 2, 'where it was');
}));

test('a table\'s resize square: a drag scales its columns and rows, one undo step; a click or Esc changes nothing', () => inQuire(PARAS(['First paragraph of text.', 'Second paragraph.']) + TABLE(1) + 'd.main.blocks[1].grid = [90, 90];', async (u) => {
  const state = () => u.evaluate(`(() => { const t = L.D.doc.main.blocks[1]; return JSON.stringify({ grid: t.grid, h: t.rows.map((r) => r.trPr.h || null) }); })()`).then(JSON.parse);
  const s0 = await state(), t0 = await u.el('#pages table.tbl');
  await u.move({ x: t0.l + 20, y: t0.t + 8 });
  let q = await u.el('.tbl-size');
  assert.ok(q && q.l >= t0.r - 1 && q.t >= t0.b - 1, 'the square shows at its bottom right corner');
  await u.click({ x: q.x, y: q.y });
  assert.deepEqual(await state(), s0, 'a click changes nothing');
  await u.move({ x: t0.l + 20, y: t0.t + 8 }); q = await u.el('.tbl-size');
  await u.press({ x: q.x, y: q.y }); await u.glide({ x: q.x, y: q.y }, { x: q.x + 60, y: q.y + 40 }); await u.esc(); await u.release({ x: q.x + 60, y: q.y + 40 });
  assert.deepEqual(await state(), s0, 'Esc: no change');
  assert.equal(await u.evaluate(`document.querySelectorAll('.tbl-ghost').length`), 0, 'no outline left');
  await u.move({ x: t0.l + 20, y: t0.t + 8 }); q = await u.el('.tbl-size');
  await u.drag({ x: q.x, y: q.y }, { x: q.x + 60, y: q.y + 40 });
  const s1 = await state(), t1 = await u.el('#pages table.tbl');
  const pt = 60 / (await u.evaluate('L.layout.zoom') * 96 / 72);
  assert.ok(Math.abs(s1.grid[0] - s1.grid[1]) < 0.1 && Math.abs(s1.grid[0] + s1.grid[1] - 180 - pt) < 1, `the columns ${pt.toFixed(1)} pt wider together, alike (${s1.grid})`);
  assert.ok(s1.h.every((h) => h > 0), `the rows given heights (${s1.h})`);
  assert.ok(Math.abs(t1.r - t0.r - 60) <= 2 && Math.abs(t1.b - t0.b - 40) <= 3, `as large as dragged (${t1.r - t0.r}, ${t1.b - t0.b})`);
  assert.deepEqual(await undoLabels(u), ['Resize Table']);
}));

test('a whole table selected by its handle: Ctrl+X cuts the table and Ctrl+V puts it back, Backspace deletes it, Delete empties its cells (Word)', async () => {
  const doc = PARAS(['First paragraph.', 'Second paragraph.']) + TABLE(1) + 'd.main.blocks[1].grid = [90, 90];';
  const select = async (u) => { const t = await u.el('#pages table.tbl'); await u.move({ x: t.l + 20, y: t.t + 8 }); const h = await u.el('.tbl-move'); await u.click({ x: h.x, y: h.y }); };
  const FULL = [[['cell 00 text', 'cell 01 text'], ['cell 10 text', 'cell 11 text']]], EMPTY = [[['', ''], ['', '']]];
  await inQuire(doc, async (u) => {
    await select(u);
    await u.key('x', { ctrl: true });
    assert.deepEqual(await body(u), ['First paragraph.', 'Second paragraph.'], 'Ctrl+X: the table is gone');
    assert.deepEqual(await undoLabels(u), ['Cut']);
    await u.key('v', { ctrl: true });
    assert.deepEqual((await body(u)).filter((b) => Array.isArray(b)), FULL, 'Ctrl+V: the table again, with its text');
  });
  await inQuire(doc, async (u) => {
    await select(u);
    await u.key('Backspace');
    assert.deepEqual(await body(u), ['First paragraph.', 'Second paragraph.'], 'Backspace: the table is gone');
    await u.key('z', { ctrl: true });
    assert.deepEqual((await body(u))[1], FULL[0], 'Undo: back, with its text');
  });
  await inQuire(doc, async (u) => {
    await select(u);
    await u.key('Delete');
    assert.deepEqual(await body(u), ['First paragraph.', EMPTY[0], 'Second paragraph.'], 'Delete: the cells are emptied, the table stays');
  });
  await inQuire(doc, async (u) => {
    const a = await u.el('#pages td[data-cid]');
    await u.drag({ x: a.l + 4, y: a.y }, { x: a.r - 4, y: a.y });
    await u.key('x', { ctrl: true });
    assert.equal((await body(u)).filter((b) => Array.isArray(b)).length, 1, 'part of it: the table stays');
  });
});

test('the task pane follows the selection: a shape opens Drawing Tools though the pane was closed, the text closes it again, a whole table shows Table Tools; its menu lists every tool panel', () => inQuire(PARAS(SHAPE_DOC) + SHAPE(1) + TABLE(1) + 'd.main.blocks[1].grid = [90, 90];', async (u) => {
  const pane = () => u.evaluate(`JSON.stringify({ open: !L.panes.task.el.hidden, current: L.panes.task.current, saved: L.app.opts.taskOpen })`).then(JSON.parse);
  if ((await pane()).open) await u.click(await u.el('#taskpane [aria-label="Close task pane"]'));
  const text = await atB(u, 0, 5);
  const shape = await u.el('#pages .flt');
  await u.click({ x: shape.x, y: shape.y });
  assert.deepEqual(await pane(), { open: true, current: 'tools-drawing', saved: false }, 'a shape: Drawing Tools, the pane left closed for next time');
  await u.click(text);
  assert.equal((await pane()).open, false, 'the text: closed again');
  const t = await u.el('#pages table.tbl');
  await u.move({ x: t.l + 20, y: t.t + 8 });
  const h = await u.el('.tbl-move');
  await u.click({ x: h.x, y: h.y });
  assert.deepEqual(await pane(), { open: true, current: 'tools-table', saved: false }, 'the whole table: Table Tools');
  await u.click(await u.el('#taskpane .tp-title'));
  const rows = await u.evaluate(`[...document.querySelectorAll('.menu .m-item')].map((r) => r.querySelector('.m-lbl').textContent + (r.classList.contains('dis') ? ' (greyed)' : ''))`);
  for (const x of ['Getting Started', 'Styles and Formatting', 'Table Tools', 'Picture Tools (greyed)', 'Drawing Tools (greyed)']) assert.ok(rows.includes(x), `the menu lists ${x}: ${rows.join(', ')}`);
  await u.esc();
  await u.click(text);
  assert.equal((await pane()).open, false);
  await u.evaluate('L.app.opts.paneFollows = false');
  await u.click({ x: shape.x, y: shape.y });
  assert.equal((await pane()).open, false, 'with the option off, the pane stays closed');
}));


test('text in a shape: a right-click on a handle offers Add Text, then typing goes in; typing on a selected shape adds to its text; a click in a box\'s empty space puts the caret in it', () => inQuire(PARAS(SHAPE_DOC) + SHAPE(1), async (u) => {
  const shapeText = () => u.evaluate(`(() => { const it = L.D.doc.main.blocks[1].runs.find((r) => r.t === 'shape'); return it.tb ? it.tb.blocks.map((p) => L.D.ptext(p)).join('/') : null; })()`);
  const inShape = () => u.evaluate(`(() => { const it = L.D.doc.main.blocks[1].runs.find((r) => r.t === 'shape'); return !!(it.tb && L.ed.sel && it.tb.blocks.includes(L.ed.sel.f.p)); })()`);
  let sh = await u.el('#pages .flt');
  await u.click({ x: sh.x, y: sh.y });
  const h = await u.el('.obj-handles .oh-se');
  await u.rightClick({ x: h.x, y: h.y });
  assert.ok((await u.menuItems()).includes('Add Text'), 'the shape\'s menu, on its handle');
  await u.pick('Add Text');
  await u.key('H'); await u.key('i');
  assert.equal(await shapeText(), 'Hi');
  assert.equal(await inShape(), true, 'the caret in its text');
  await u.click(await atB(u, 0, 5));
  sh = await u.el('#pages .flt');
  await u.click({ x: sh.l + 2, y: sh.y });   // its border selects it
  assert.equal(await u.evaluate('!!L.ed.objSel'), true, 'setup: the shape is selected');
  await u.key('!'); await u.key('?');
  assert.equal(await shapeText(), 'Hi!?', 'typing on the selected shape goes to the end of its text');
  assert.deepEqual(await body(u), SHAPE_DOC.map((s, i) => (i === 1 ? '\ufffc' : '') + s), 'nothing typed into the page');
  await u.click(await atB(u, 0, 5));
  const area = await u.el('#pages .flt .tbx');
  await u.click({ x: area.x, y: area.b - 3 });   // its text area, below the line
  assert.equal(await inShape(), true, 'a click in the empty part of its text area puts the caret in its text');
  assert.deepEqual(await undoLabels(u), ['Add Text', 'Typing', 'Typing'], 'one step for each run of typing');
}));

test('the toolbar\'s Insert Table grid: a move over 2 × 3 and a click insert that table', () => inQuire(PARAS(['First paragraph.', 'Second paragraph.']), async (u) => {
  await u.click(await u.el('.tb-drop[data-tip="Insert Table"]'));
  const cell = (k) => u.el(`.tgp-cells span:nth-child(${k})`);
  const a = await cell(1);
  await u.glide(a, await cell(8), 6);   // row 2, column 3 of the 5-wide grid
  assert.equal(await u.evaluate(`document.querySelector('.tgp-label').textContent`), '2 x 3 Table');
  await u.click(await cell(8));
  assert.equal(await u.evaluate(`!!document.querySelector('.tgp-cells')`), false, 'the menu closes');
  const tb = await u.evaluate(`L.D.doc.main.blocks.filter((b) => b.t === 'tbl').map((t) => [t.rows.length, t.rows[0].cells.length])`);
  assert.deepEqual(tb, [[2, 3]]);
}));

test('PDF export draws a table as on screen: one line between cells, the columns as wide', () => inQuire(PARAS(['First paragraph.', 'Second paragraph.']) + 'd.main.blocks.splice(1, 0, D.simpleTable(2, 2, 432));', async (u) => {
  // the page picture in the PDF, sampled along the middle of the table's first row: where its dark lines are
  const r = await u.evaluate(`(async () => {
    const pg = L.layout.pages[0].el, P = pg.getBoundingClientRect(), tds = [...pg.querySelectorAll('td')].slice(0, 2).map((e) => e.getBoundingClientRect());
    const bytes = new Uint8Array(await (await L.app.buildPDF()).arrayBuffer()); L.ui.busy(false);
    let s = 0; while (!(bytes[s] === 0xff && bytes[s + 1] === 0xd8)) s++;
    let e = bytes.length; while (!(bytes[e - 2] === 0xff && bytes[e - 1] === 0xd9)) e--;
    const img = await createImageBitmap(new Blob([bytes.subarray(s, e)], { type: 'image/jpeg' })), k = img.width / P.width;
    const c = new OffscreenCanvas(img.width, img.height), g = c.getContext('2d'); g.drawImage(img, 0, 0);
    const y = Math.round((tds[0].top + tds[0].bottom) / 2 - P.top) * k, x0 = Math.round((tds[0].left - P.left - 6) * k), x1 = Math.round((tds[1].right - P.left + 6) * k);
    const px = g.getImageData(x0, Math.round(y), x1 - x0, 1).data, lines = [];
    for (let x = 0, on = false; x < x1 - x0; x++) { const dark = px[x * 4] < 128; if (dark && !on) lines.push(Math.round((x + x0) / k - (tds[0].left - P.left))); on = dark; }
    return { lines, w: [tds[0].width, tds[1].width].map(Math.round) };
  })()`);
  assert.equal(r.lines.length, 3, `three lines across two cells (${r.lines}), not a box round each cell`);
  assert.ok(Math.abs(r.lines[1] - r.w[0]) <= 2 && Math.abs(r.lines[2] - r.w[0] - r.w[1]) <= 2, `the lines where the screen has them (${r.lines}; cells ${r.w})`);
}));

test('AutoCorrect while typing: a typo, a contraction without its apostrophe, i, a loanword and (c)', () => inQuire(PARAS(['']), async (u) => {
  await u.type('so i dont think teh cafe is (c) ');
  assert.equal(await u.evaluate('L.D.ptext(L.D.doc.main.blocks[0])'), "So I don't think the café is © ");
}));
