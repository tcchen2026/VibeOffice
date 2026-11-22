// Table properties belong to the cell/table, including through text edits and history.
// BASE=... CDP=... node --test tests/lectern/tables.test.mjs
import { after, before, test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { ensureServer, openPage } from '../../tools/ooxml/cdp.mjs';
let page, stop;
before(async () => { stop = await ensureServer(); page = await openPage('lectern', { persistence: false, automated: true }); });
after(async () => { try { await page?.close(); } finally { stop?.(); } });
async function run(fn, arg) {
  const result = await page.evaluate(`(${fn.toString()})(${JSON.stringify(arg)})`);
  assert.deepEqual(page.errors, [], 'no browser errors');
  return result;
}

test('tcPr text direction retains all Office values, absence, text edits, copy, history and drafts', async () => {
  const r = await run(async () => {
    const dirs = [null, 'horz', 'vert', 'vert270', 'wordArtVert', 'eaVert', 'mongolianVert', 'wordArtVertRtl'];
    const p = L.model.newPresentation(), sl = p.slides[0];
    sl.shapes = [L.model.newTable(p, sl, 1, dirs.length, 20, 20, 720)];
    sl.shapes[0].tbl.rows[0].h = 130;
    for (const c of sl.shapes[0].tbl.rows[0].cells) c.tx.ps = [L.txt.para('Turn me')];
    const z = await L.zip.read(new Uint8Array(await (await L.pptx.write(p)).arrayBuffer())), entries = [];
    // Author the input attributes independently of the writer under test.
    for (const item of new Set(z.values())) {
      let data = await item.bytes();
      if (item.name === 'ppt/slides/slide1.xml') {
        let i = 0;
        data = (await item.text()).replace(/<a:tcPr\b[^>]*>/g, m => { const dir = dirs[i++]; return m.replace(/\svert="[^"]*"/, '').replace(/>$/, (dir ? ' vert="' + dir + '"' : '') + '>'); });
      }
      entries.push({ name: item.name, data });
    }
    L.pres = await L.pptx.read(new Uint8Array(await (await L.zip.write(entries)).arrayBuffer()));
    L.app.loadPres(L.pres, 'Directions', { saved: true }); L.hist.clear();
    const table = () => L.pres.slides[0].shapes.find(s => s.tbl);
    const values = () => table().tbl.rows[0].cells.map(c => c.tx.vert ?? null);
    const saved = async () => {
      const pkg = await L.opc.open(new Uint8Array(await (await L.pptx.write(L.pres)).arrayBuffer()));
      return [...pkg.xml('ppt/slides/slide1.xml').getElementsByTagNameNS(L.opc.NS.a, 'tcPr')].map(e => e.getAttribute('vert'));
    };
    const imported = values(), unedited = await saved();
    const el = L.render.slide(L.pres, L.pres.slides[0]);
    const transforms = [...el.querySelectorAll('td')].map(td => td.querySelector('.tx')?.style.transform || '');
    L.hist.push('Cell text'); table().tbl.rows[0].cells[0].tx.ps[0].rs[0].t += '!';
    const edited = await saved();
    L.hist.doUndo(); const undo = values(); L.hist.doRedo(); const redo = values();
    const copied = L.model.dup(table()).tbl.rows[0].cells.map(c => c.tx.vert ?? null);
    const clip = L.preserve.pasteboard(JSON.parse(JSON.stringify(L.preserve.clipboard({ kind: 'shapes', shapes: [L.clone(table())], text: '' }))));
    const pasted = clip.shapes[0].tbl.rows[0].cells.map(c => c.tx.vert ?? null);
    const blob = await __corpusHooks.snapshot(L.pres);
    L.hist.clear();
    await __corpusHooks.open(new File([blob], 'Directions.pptx'), { draft: true, name: 'Directions.pptx', saved: true });
    const draft = values();
    table().tbl.rows[0].cells[2].tx.vert = 'horz';
    const replaced = await saved();
    delete table().tbl.rows[0].cells[2].tx.vert;
    return { imported, unedited, edited, undo, redo, copied, pasted, draft, replaced, removed: await saved(), transforms };
  });
  const expected = [null, 'horz', 'vert', 'vert270', 'wordArtVert', 'eaVert', 'mongolianVert', 'wordArtVertRtl'];
  for (const k of ['imported', 'unedited', 'edited', 'undo', 'redo', 'copied', 'pasted', 'draft']) assert.deepEqual(r[k], expected, k);
  assert.equal(r.transforms[2], 'rotate(90deg)'); assert.equal(r.transforms[3], 'rotate(-90deg)');
  assert.deepEqual(r.replaced, expected.map((v, i) => i === 2 ? 'horz' : v));
  assert.deepEqual(r.removed, expected.map((v, i) => i === 2 ? null : v));
});

test('RTL table columns keep logical cell identities and Tab follows reading order', async () => {
  const r = await run(async () => {
    const p = L.model.newPresentation(), sl = p.slides[0];
    sl.shapes = [L.model.newTable(p, sl, 2, 3, 50, 60, 540)];
    sl.shapes[0].tbl.rows.forEach((row, r) => { row.h = 80; row.cells.forEach((c, i) => c.tx.ps = [L.txt.para(`${r}:${i}`)]); });
    const z = await L.zip.read(new Uint8Array(await (await L.pptx.write(p)).arrayBuffer())), entries = [];
    for (const item of new Set(z.values())) entries.push({ name: item.name, data: item.name === 'ppt/slides/slide1.xml' ? (await item.text()).replace('<a:tblPr ', '<a:tblPr rtl="1" ') : await item.bytes() });
    const pres = await L.pptx.read(new Uint8Array(await (await L.zip.write(entries)).arrayBuffer()));
    L.app.loadPres(pres, 'RTL', { saved: true }); L.hist.clear();
    const table = () => L.pres.slides[0].shapes.find(s => s.tbl);
    const saved = async () => {
      const pkg = await L.opc.open(new Uint8Array(await (await L.pptx.write(L.pres)).arrayBuffer()));
      return pkg.xml('ppt/slides/slide1.xml').getElementsByTagNameNS(L.opc.NS.a, 'tblPr')[0].getAttribute('rtl');
    };
    const imported = table().tbl.rtl ?? null, unedited = await saved();
    L.hist.push('Text'); table().tbl.rows[0].cells[0].tx.ps[0].rs[0].t += '!';
    const edited = await saved(); L.hist.doUndo(); const undo = table().tbl.rtl; L.hist.doRedo(); const redo = table().tbl.rtl;
    const copied = L.model.dup(table()).tbl.rtl;
    const pasted = L.preserve.pasteboard(JSON.parse(JSON.stringify(L.preserve.clipboard({ kind: 'shapes', shapes: [L.clone(table())], text: '' })))).shapes[0].tbl.rtl;
    const blob = await __corpusHooks.snapshot(L.pres); L.hist.clear();
    await __corpusHooks.open(new File([blob], 'RTL.pptx'), { draft: true, name: 'RTL.pptx', saved: true });
    const draft = table().tbl.rtl; table().tbl.rtl = false; const changed = await saved();
    delete table().tbl.rtl; const removed = await saved(); table().tbl.rtl = true;
    L.ed.goto(0, { force: true }); L.ed.render();
    return { imported, unedited, edited, undo, redo, copied, pasted, draft, changed, removed };
  });
  assert.equal(r.imported, true); assert.equal(r.unedited, '1'); assert.equal(r.edited, '1');
  for (const k of ['undo', 'redo', 'copied', 'pasted', 'draft']) assert.equal(r[k], true, k);
  assert.equal(r.changed, '0'); assert.equal(r.removed, null);
  const cells = await page.evaluate(`Array.from(document.querySelectorAll('.sl-edit td')).map(td => { const r = td.getBoundingClientRect(); return { r:+td.dataset.r,c:+td.dataset.c,x:r.x+r.width/2,y:r.y+r.height/2 }; })`);
  assert.equal(cells.length, 6); assert.ok(cells[0].x > cells[1].x && cells[1].x > cells[2].x);
  const pause = () => new Promise(resolve => setTimeout(resolve, 150));
  const click = async c => { await page.send('Input.dispatchMouseEvent', { type: 'mousePressed', x:c.x, y:c.y, button:'left', clickCount:1 }); await page.send('Input.dispatchMouseEvent', { type:'mouseReleased', x:c.x, y:c.y, button:'left', clickCount:1 }); await pause(); };
  await click(cells[0]);
  assert.deepEqual(await page.evaluate('L.te.state.cell'), [0,0]);
  await page.send('Input.dispatchKeyEvent', { type:'keyDown', key:'Tab', code:'Tab', windowsVirtualKeyCode:9 });
  await page.send('Input.dispatchKeyEvent', { type:'keyUp', key:'Tab', code:'Tab', windowsVirtualKeyCode:9 }); await pause();
  assert.deepEqual(await page.evaluate('L.te.state.cell'), [0,1], 'Tab moves visually left to the next logical cell');
  await click(cells[0]);
  await page.send('Input.dispatchMouseEvent', { type:'mousePressed', x:cells[0].x, y:cells[0].y, button:'left', clickCount:1 }); await pause();
  await page.send('Input.dispatchMouseEvent', { type:'mouseMoved', x:cells[4].x, y:cells[4].y, button:'left', buttons:1 }); await pause();
  await page.send('Input.dispatchMouseEvent', { type:'mouseReleased', x:cells[4].x, y:cells[4].y, button:'left', clickCount:1 }); await pause();
  assert.deepEqual(await page.evaluate('L.ed.cellSel'), { r0:0,c0:0,r1:1,c1:1 });
});

test('table theme fills and borders retain transforms through independent edits', async () => {
  const r = await run(async () => {
    const p = L.model.newPresentation(), sl = p.slides[0];
    sl.shapes = [L.model.newTable(p, sl, 1, 1, 20, 30, 240)];
    const z = await L.zip.read(new Uint8Array(await (await L.pptx.write(p)).arrayBuffer())), entries = [];
    const theme = '<a:schemeClr val="accent2"><a:lumMod val="75000"/><a:lumOff val="10000"/></a:schemeClr>';
    for (const item of new Set(z.values())) entries.push({ name:item.name, data:item.name === 'ppt/slides/slide1.xml' ? (await item.text()).replace(/<a:tcPr\b[^>]*>.*?<\/a:tcPr>/gs, '<a:tcPr><a:lnL w="25400"><a:solidFill>'+theme+'</a:solidFill><a:prstDash val="solid"/></a:lnL><a:solidFill>'+theme+'</a:solidFill></a:tcPr>') : await item.bytes() });
    L.app.loadPres(await L.pptx.read(new Uint8Array(await (await L.zip.write(entries)).arrayBuffer())), 'Theme', { saved:true }); L.hist.clear();
    const cell = () => L.pres.slides[0].shapes[0].tbl.rows[0].cells[0];
    const saved = async () => {
      const pkg = await L.opc.open(new Uint8Array(await (await L.pptx.write(L.pres)).arrayBuffer()));
      const pr = pkg.xml('ppt/slides/slide1.xml').getElementsByTagNameNS(L.opc.NS.a,'tcPr')[0];
      const direct = [...pr.children].find(e=>e.localName==='solidFill'), border = [...pr.children].find(e=>e.localName==='lnL');
      const vals = e => [...e.getElementsByTagNameNS(L.opc.NS.a,'schemeClr')].map(c=>[c.getAttribute('val'),...c.children.map(x=>[x.localName,x.getAttribute('val')])]);
      return { fill:vals(direct), line:vals(border), width:border.getAttribute('w'), rgb:[...border.getElementsByTagNameNS(L.opc.NS.a,'srgbClr')].map(e=>e.getAttribute('val')) };
    };
    const original = await saved(); L.hist.push('Border width'); cell().bd.l.w=3;
    const width = await saved(); L.hist.doUndo(); const undo = await saved(); L.hist.doRedo(); const redo = await saved();
    cell().tx.ps[0].rs=[{t:'Changed'}];
    L.model.design(L.pres,L.pres.slides[0]).colors.accent2='#336699';
    const themeChanged=await saved();
    const paint=L.properties.cellPaint?.(cell(),L.model.design(L.pres,L.pres.slides[0]));
    const expectedPaint=L.color.applyMods('#336699',[{name:'lumMod',val:75000},{name:'lumOff',val:10000}]);
    const copy=L.preserve.pasteboard(JSON.parse(JSON.stringify(L.preserve.clipboard({kind:'shapes',shapes:[L.model.dup(L.pres.slides[0].shapes[0])],text:''}))));
    L.pres.slides[0].shapes=copy.shapes; const copied=await saved();
    const blob=await __corpusHooks.snapshot(L.pres);L.hist.clear();
    await __corpusHooks.open(new File([blob],'Theme.pptx'),{draft:true,name:'Theme.pptx',saved:true});const draft=await saved();
    cell().bd.l.c='#123456'; const rgb=await saved();
    return {original,width,undo,redo,themeChanged,copied,draft,rgb,paint,expectedPaint};
  });
  const token=[['accent2',['lumMod','75000'],['lumOff','10000']]];
  for(const key of ['original','width','undo','redo','themeChanged','copied','draft']) { assert.deepEqual(r[key].fill,token,key+' fill');assert.deepEqual(r[key].line,token,key+' border'); }
  assert.equal(r.paint.fill.c,r.expectedPaint);assert.equal(r.paint.bd.l.c,r.expectedPaint);
  assert.equal(r.original.width,'25400');assert.equal(r.width.width,'38100');assert.equal(r.undo.width,'25400');assert.equal(r.redo.width,'38100');
  assert.deepEqual(r.rgb.line,[]);assert.deepEqual(r.rgb.rgb,['123456']);assert.deepEqual(r.rgb.fill,token);
});

test('table style link and sparse inherited formatting survive text edits and detached copies', async t => {
  const file = path.join(process.env.VO_CORPORA || '/work/corpora', 'libreoffice/sd/qa/unit/data/pptx/tcPr-vert-roundtrip.pptx');
  if (!fs.existsSync(file)) return t.skip('Fetch pinned corpus: '+file);
  const r=await run(async data=>{
    const bytes=Uint8Array.from(atob(data),c=>c.charCodeAt(0));
    L.app.loadPres(await L.pptx.read(bytes),'Style',{saved:true});L.hist.clear();
    const table=()=>L.pres.slides[0].shapes.find(s=>s.tbl), cell=()=>table().tbl.rows[0].cells[0];
    const saved=async()=>{
      const bytes=new Uint8Array(await (await L.pptx.write(L.pres)).arrayBuffer()), pkg=await L.opc.open(bytes);
      let binary=''; for(let i=0;i<bytes.length;i+=32768) binary+=String.fromCharCode(...bytes.subarray(i,i+32768));
      const xml=pkg.xml('ppt/slides/slide1.xml'), stylePart=pkg.rels(pkg.main).find(r=>/\/tableStyles$/.test(r.type)).part;
      return {ids:[...xml.getElementsByTagNameNS(L.opc.NS.a,'tableStyleId')].map(e=>e.textContent),
        paint:[...xml.getElementsByTagNameNS(L.opc.NS.a,'tcPr')].map(e=>e.children.map(c=>c.localName)),
        runs:[...xml.getElementsByTagNameNS(L.opc.NS.a,'rPr')].map(e=>({b:e.getAttribute('b'),fill:[...e.getElementsByTagNameNS(L.opc.NS.a,'solidFill')].map(L.opc.raw)})),
        styles:pkg.text(stylePart), bytes:btoa(binary), pkg};
    };
    const compact=v=>{const {pkg,...rest}=v;return rest;};
    const initial=await saved(), styleId=table().tbl.styleId;
    L.hist.push('Cell text');cell().tx.ps[0].rs[0].t+='!';const edited=await saved();
    L.hist.doUndo();const undo=await saved();L.hist.doRedo();const redo=await saved();
    const packet=JSON.parse(JSON.stringify(L.preserve.clipboard({kind:'shapes',shapes:[L.model.dup(table())],text:''})));
    const blob=await __corpusHooks.snapshot(L.pres);L.hist.clear();
    await __corpusHooks.open(new File([blob],'Style.pptx'),{draft:true,name:'Style.pptx',saved:true});const draft=await saved();
    cell().tx.ps[0].rs[0].color='#881133';const explicit=await saved();
    // A detached copy in a presentation without this style gets its definition as well as its link.
    L.pres=L.model.newPresentation();L.pres.slides[0].shapes=L.preserve.pasteboard(packet).shapes;
    const copied=await saved();
    // A destination with a conflicting custom definition keeps its original; the imported one gets a UUID.
    const z=await L.zip.read(bytes), entries=[];
    for(const item of new Set(z.values())) entries.push({name:item.name,data:item.name==='ppt/tableStyles.xml'?(await item.text()).replace(/b="on"/g,'b="off"'):await item.bytes()});
    L.pres=await L.pptx.read(new Uint8Array(await (await L.zip.write(entries)).arrayBuffer()));
    L.pres.slides[0].shapes.push(...L.preserve.pasteboard(packet).shapes);
    const collision=await saved(), repeat=await saved();
    return {styleId,initial:compact(initial),edited:compact(edited),undo:compact(undo),redo:compact(redo),draft:compact(draft),explicit:compact(explicit),copied:compact(copied),collision:compact(collision),repeat:compact(repeat)};
  },fs.readFileSync(file).toString('base64'));
  if(process.env.TABLE_ARTIFACTS) {
    fs.mkdirSync(process.env.TABLE_ARTIFACTS,{recursive:true});
    for(const k of ['edited','copied','collision']) fs.writeFileSync(path.join(process.env.TABLE_ARTIFACTS,k+'.pptx'),Buffer.from(r[k].bytes,'base64'));
  }
  const id='{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}';assert.equal(r.styleId,id);
  for(const k of ['initial','edited','undo','redo','draft','copied']) {
    assert.deepEqual(r[k].ids,[id],k);assert.ok(r[k].paint.every(p=>!p.length),k+' inherits paint');
    assert.ok(r[k].runs.every(p=>p.b===null&&!p.fill.length),k+' inherits text');
    assert.match(r[k].styles,/styleId="\{5C22544A-7EE6-4342-B048-85BDC9FD1C3A\}"/);
  }
  assert.match(r.explicit.runs[0].fill[0],/881133/);assert.equal(r.explicit.runs[0].b,null);
  assert.match(r.copied.styles,/\bdef="\{5C22544A-7EE6-4342-B048-85BDC9FD1C3A\}"/,'a new table-style list names the built-in default');
  assert.equal(r.collision.ids[0],id);assert.notEqual(r.collision.ids[1],id);
  assert.match(r.collision.ids[1],/^\{[0-9A-F]{8}-[0-9A-F]{4}-4[0-9A-F]{3}-[89AB][0-9A-F]{3}-[0-9A-F]{12}\}$/);
  assert.ok(r.collision.styles.includes('styleId="'+r.collision.ids[1]+'"'));
  assert.deepEqual(r.repeat.ids,r.collision.ids,'the remapped copy id is stable across saves');
});

test('table geometry keeps original EMUs and an edit owns only its dimension', async () => {
  const r=await run(async()=>{
    const p=L.model.newPresentation(),sl=p.slides[0];sl.shapes=[L.model.newTable(p,sl,2,2,20,30,500)];
    const z=await L.zip.read(new Uint8Array(await (await L.pptx.write(p)).arrayBuffer())),entries=[];
    for(const item of new Set(z.values())) {
      let data=await item.bytes();
      if(item.name==='ppt/slides/slide1.xml') {
        let c=0,r=0,i=0;
        data=(await item.text()).replace(/<a:gridCol w="\d+"/g,()=>'<a:gridCol w="'+[5504878,914401][c++]+'"')
          .replace(/<a:tr h="\d+"/g,()=>'<a:tr h="'+[1066801,1066799][r++]+'"')
          .replace(/<a:tcPr\b[^>]*>/g,m=>{m=m.replace(/ mar[LTRB]="[^"]*"/g,'');return i++?m:m.replace('>',' marL="91444" marT="45721" marR="91444" marB="45721">');});
      }
      entries.push({name:item.name,data});
    }
    L.app.loadPres(await L.pptx.read(new Uint8Array(await (await L.zip.write(entries)).arrayBuffer())),'Geometry',{saved:true});L.hist.clear();
    const table=()=>L.pres.slides[0].shapes[0].tbl;
    const saved=async()=>{
      const pkg=await L.opc.open(new Uint8Array(await (await L.pptx.write(L.pres)).arrayBuffer())),xml=pkg.xml('ppt/slides/slide1.xml');
      const els=name=>[...xml.getElementsByTagNameNS(L.opc.NS.a,name)];
      return {cols:els('gridCol').map(e=>+e.getAttribute('w')),rows:els('tr').map(e=>+e.getAttribute('h')),
        margins:els('tcPr').map(e=>['marL','marT','marR','marB'].map(n=>e.getAttribute(n)))};
    };
    const original=await saved();table().rows[0].cells[0].tx.ps[0].rs=[{t:'Edited text'}];const text=await saved();
    L.hist.push('Table dimensions');table().cols[1]+=1.25;table().rows[0].h+=0.5;table().rows[0].cells[0].tx.ins[2]+=2;
    const changed=await saved();L.hist.doUndo();const undo=await saved();L.hist.doRedo();const redo=await saved();
    const copied=L.preserve.pasteboard(JSON.parse(JSON.stringify(L.preserve.clipboard({kind:'shapes',shapes:[L.model.dup(L.pres.slides[0].shapes[0])],text:''}))));
    L.pres.slides[0].shapes=copied.shapes;const copy=await saved();
    const blob=await __corpusHooks.snapshot(L.pres);L.hist.clear();
    await __corpusHooks.open(new File([blob],'Geometry.pptx'),{draft:true,name:'Geometry.pptx',saved:true});const draft=await saved();
    return {original,text,changed,undo,redo,copy,draft};
  });
  const original={cols:[5504878,914401],rows:[1066801,1066799],margins:[['91444','45721','91444','45721'],...Array.from({length:3},()=>[null,null,null,null])]};
  for(const k of ['original','text','undo']) assert.deepEqual(r[k],original,k);
  const changed=structuredClone(original);changed.cols[1]+=15875;changed.rows[0]+=6350;changed.margins[0][2]='116844';
  for(const k of ['changed','redo','copy','draft']) assert.deepEqual(r[k],changed,k);
});

// Text turned 90° in a cell runs along the cell as tall as its rows grew: POI's table_test2 stores the rows at less
// than their text needs, and PowerPoint 2016 wraps "erat, sed diam voluptua. At vero eos et accusam et justo" into
// five lines along the grown cell (2026-11-07); Lectern wrapped it into ten along the stored height.
test('rotated cell text wraps along the height the cell grew to', async t => {
  const file = path.join(process.env.VO_CORPORA || '/work/corpora', 'powerpoint', 'poi__f01b52f70acc__table_test2.pptx');
  if (!fs.existsSync(file)) return t.skip('Fetch pinned corpus: ' + file);
  const lines = await run(async data => {
    const pres = await L.pptx.read(Uint8Array.from(atob(data), c => c.charCodeAt(0)));
    const el = L.render.slide(pres, pres.slides[0], { mode: 'thumb', index: 0 });
    el.style.cssText += ';position:fixed;left:0;top:0';
    document.body.appendChild(el);
    await document.fonts.ready;
    const p = [...el.querySelectorAll('td p.pa')].find(x => x.textContent.startsWith('erat'));
    const range = document.createRange(); range.selectNodeContents(p);
    // turned 90°, each line is a column: count the columns the words stand in
    const n = new Set([...range.getClientRects()].map(r => Math.round(r.left))).size;
    el.remove();
    return n;
  }, fs.readFileSync(file).toString('base64'));
  assert.equal(lines, 5);
});
