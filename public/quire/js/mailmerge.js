/* Quire — mail merge: data sources (.csv/.txt/.docx tables or a typed list), recipient filtering,
 * address block / greeting line / merge fields, preview, and merging to a new document or PDF.
 */
(function () {
  'use strict';
  const L = window.L, D = L.D, O = L.O, E = L.ed, LY = L.layout;
  const { h } = L;
  const ui = L.ui;
  const MM = (L.mailmerge = { data: null, rec: 0, preview: false, type: 'letters', wstep: 1 });
  const doc = () => D.doc;
  const A = () => L.app;
  const norm = (s) => String(s).toLowerCase().replace(/[^a-z0-9]/g, '');
  const ADDRESS_FIELDS = ['Title', 'First Name', 'Last Name', 'Company Name', 'Address Line 1', 'Address Line 2', 'City', 'State', 'ZIP Code', 'Country', 'Home Phone', 'Work Phone', 'E-mail Address'];

  /* ---------- data sources ---------- */
  function parseDelimited(text) {
    text = text.replace(/^﻿/, '');
    const first = text.split(/\r?\n/)[0] || '';
    const delim = (first.match(/\t/g) || []).length > (first.match(/,/g) || []).length ? '\t' : (first.match(/;/g) || []).length > (first.match(/,/g) || []).length ? ';' : ',';
    const rows = [];
    let row = [], cell = '', q = false;
    for (let i = 0; i < text.length; i++) {
      const c = text[i];
      if (q) { if (c === '"') { if (text[i + 1] === '"') { cell += '"'; i++; } else q = false; } else cell += c; continue; }
      if (c === '"' && !cell) { q = true; continue; }
      if (c === delim) { row.push(cell); cell = ''; continue; }
      if (c === '\n' || c === '\r') { if (c === '\r' && text[i + 1] === '\n') i++; row.push(cell); cell = ''; if (row.some((x) => x !== '')) rows.push(row); row = []; continue; }
      cell += c;
    }
    row.push(cell);
    if (row.some((x) => x !== '')) rows.push(row);
    return rows;
  }
  function setData(name, rows) {
    if (!rows.length) { ui.msg('The data source contains no records.', { icon: 'warn' }); return false; }
    const fields = rows[0].map((f, i) => String(f).trim() || 'Field' + (i + 1));
    MM.data = { name, fields, rows: rows.slice(1).map((r) => fields.map((_, i) => (r[i] == null ? '' : String(r[i])))), included: null };
    MM.data.included = new Set(MM.data.rows.map((_, i) => i));
    MM.rec = 0;
    A().opts.toolbars.mailMerge = true;
    A().updateToolbars();
    refreshPane();
    return true;
  }
  MM.openSource = async function () {
    const files = await L.pickFiles('.csv,.txt,.tsv,.docx,.htm,.html');
    const f = files[0];
    if (!f) return;
    try {
      if (/\.docx$/i.test(f.name)) {
        const res = await L.docx.read(await L.readAsArrayBuffer(f));
        const t = res.doc.main.blocks.find((b) => b.t === 'tbl');
        let rows;
        if (t) rows = t.rows.map((r) => r.cells.map((c) => c.blocks.filter((b) => b.t === 'p').map((p) => D.plainText(p)).join(' ').trim()));
        else rows = res.doc.main.blocks.filter((b) => b.t === 'p' && D.plainText(b).trim()).map((p) => D.plainText(p).split('\t'));
        setData(f.name, rows);
      } else if (/\.html?$/i.test(f.name)) {
        const dom = new DOMParser().parseFromString(await L.readAsText(f), 'text/html');
        const t = dom.querySelector('table');
        if (!t) { ui.msg('No table was found in that file.', { icon: 'warn' }); return; }
        setData(f.name, Array.from(t.rows).map((r) => Array.from(r.cells).map((c) => c.textContent.trim())));
      } else setData(f.name, parseDelimited(await L.readAsText(f)));
      MM.recipientsDialog();
    } catch (e) { ui.msg('The data source could not be opened. ' + (e.message || ''), { icon: 'error' }); }
  };
  /** Type a New Address List */
  MM.typeList = function () {
    const fields = ADDRESS_FIELDS.slice();
    const rows = MM.data && MM.data.typed ? MM.data.rows.map((r) => r.slice()) : [fields.map(() => '')];
    let cur = 0;
    const inputs = fields.map((f, i) => { const inp = h('input', { type: 'text', id: 'mm-f' + i, style: 'width:100%' }); inp.addEventListener('input', () => { rows[cur][i] = inp.value; }); return inp; });
    const grid = h('div', { style: 'display:grid;grid-template-columns:120px 1fr;gap:3px 8px;max-height:300px;overflow:auto;padding-right:4px' });
    fields.forEach((f, i) => grid.append(h('label', { for: 'mm-f' + i, text: f }), inputs[i]));
    const info = h('span');
    const load = () => { inputs.forEach((inp, i) => { inp.value = rows[cur][i] || ''; }); info.textContent = `Entry ${cur + 1} of ${rows.length}`; };
    const nav = h('div', { class: 'tp-btns' },
      ui.button('&New Entry', () => { rows.push(fields.map(() => '')); cur = rows.length - 1; load(); inputs[0].focus(); }, { class: 'btn small' }),
      ui.button('&Delete Entry', () => { if (rows.length > 1) { rows.splice(cur, 1); cur = Math.min(cur, rows.length - 1); load(); } }, { class: 'btn small' }),
      ui.button('<', () => { if (cur > 0) { cur--; load(); } }, { class: 'btn small' }), ui.button('>', () => { if (cur < rows.length - 1) { cur++; load(); } }, { class: 'btn small' }), info);
    load();
    ui.dialog({
      title: 'New Address List', width: 440, body: h('div', { class: 'col' }, grid, nav), buttons: [{ label: 'OK', primary: true, onClick: () => { const data = [fields].concat(rows.filter((r) => r.some(Boolean))); if (setData('Address list', data)) MM.data.typed = true; } }, { label: 'Cancel' }],
    });
  };
  MM.recipientsDialog = function () {
    const d = MM.data;
    if (!d) return;
    let sortCol = -1, desc = false;
    const box = h('div', { class: 'datasheet', style: 'max-height:280px' });
    const draw = () => {
      L.clear(box);
      const order = d.rows.map((_, i) => i);
      if (sortCol >= 0) order.sort((a, b) => { const r = String(d.rows[a][sortCol]).localeCompare(String(d.rows[b][sortCol]), undefined, { numeric: true }); return desc ? -r : r; });
      const t = h('table', { class: 'mm-table' });
      const hr = h('tr', null, h('th', null, (() => { const c = h('input', { type: 'checkbox', 'aria-label': 'Include all' }); c.checked = d.included.size === d.rows.length; c.addEventListener('change', () => { d.included = c.checked ? new Set(d.rows.map((_, i) => i)) : new Set(); draw(); }); return c; })()));
      d.fields.forEach((f, i) => { const th = h('th', { text: f + (sortCol === i ? (desc ? ' ▼' : ' ▲') : ''), style: 'cursor:pointer' }); th.addEventListener('click', () => { if (sortCol === i) desc = !desc; else { sortCol = i; desc = false; } draw(); }); hr.appendChild(th); });
      t.appendChild(hr);
      for (const i of order) {
        const c = h('input', { type: 'checkbox', 'aria-label': 'Include' });
        c.checked = d.included.has(i);
        c.addEventListener('change', () => { if (c.checked) d.included.add(i); else d.included.delete(i); });
        t.appendChild(h('tr', null, h('td', null, c), ...d.rows[i].map((v) => h('td', { text: v }))));
      }
      box.appendChild(t);
    };
    draw();
    ui.dialog({ title: 'Mail Merge Recipients', width: 620, body: h('div', { class: 'col' }, h('div', { text: `Data source: ${d.name}. Click a column heading to sort; clear a check box to leave a recipient out.` }), box), buttons: [{ label: 'OK', primary: true, onClick: () => { MM.order = null; refreshPane(); } }] });
  };
  MM.records = () => (MM.data ? MM.data.rows.map((r, i) => i).filter((i) => MM.data.included.has(i)) : []);
  MM.value = function (name) {
    const d = MM.data;
    if (!d) return '';
    const recs = MM.records();
    const row = d.rows[recs[L.clamp(MM.rec, 0, recs.length - 1)]];
    if (!row) return '';
    const k = norm(name);
    let i = d.fields.findIndex((f) => norm(f) === k);
    if (i < 0) { const alias = { firstname: ['first', 'fname', 'givenname'], lastname: ['last', 'lname', 'surname', 'familyname'], addressline1: ['address', 'address1', 'street'], zipcode: ['zip', 'postalcode', 'postcode'], state: ['province', 'region'], companyname: ['company', 'organization'] }[k] || []; i = d.fields.findIndex((f) => alias.includes(norm(f))); }
    return i >= 0 ? row[i] : '';
  };
  const fieldFor = (want) => { const d = MM.data; if (!d) return null; const k = norm(want); const f = d.fields.find((x) => norm(x) === k) || d.fields.find((x) => norm(x).includes(k)); return f || null; };
  const mf = (name) => name.replace(/\s+/g, '_');

  /* ---------- inserting fields ---------- */
  function insertParts(parts, label) {
    E.edit(label, () => {
      let pos = E.deleteSelection();
      for (const part of parts) {
        if (part === '\n') { pos = O.splitPara(pos); continue; }
        if (typeof part === 'string') { pos = O.insertText(pos, part, O.inheritRPr(pos)); continue; }
        const instr = `MERGEFIELD ${mf(part.f)}`;
        pos = O.insertField(pos, instr, L.fields.evaluate(instr, pos), O.inheritRPr(pos));
      }
      return pos;
    });
  }
  MM.insertField = (name) => insertParts([{ f: name }], 'Insert Merge Field');
  MM.insertBlock = function (kind) {
    if (!MM.data) return;
    if (kind === 'greeting') {
      const parts = ['Dear '];
      const t = fieldFor('Title'), fn = fieldFor('First Name'), ln = fieldFor('Last Name');
      if (t) parts.push({ f: t }, ' ');
      if (fn) parts.push({ f: fn }, ' ');
      if (ln) parts.push({ f: ln });
      parts.push(',');
      insertParts(parts, 'Insert Greeting Line');
      return;
    }
    const parts = [];
    const add = (names, sep) => { const fs = names.map(fieldFor).filter(Boolean); if (!fs.length) return false; fs.forEach((f, i) => { if (i) parts.push(sep || ' '); parts.push({ f }); }); return true; };
    if (add(['Title', 'First Name', 'Last Name'])) parts.push('\n');
    if (add(['Company Name'])) parts.push('\n');
    if (add(['Address Line 1'])) parts.push('\n');
    if (add(['Address Line 2'])) parts.push('\n');
    const city = fieldFor('City'), st = fieldFor('State'), zip = fieldFor('ZIP Code');
    if (city) { parts.push({ f: city }); if (st || zip) parts.push(', '); }
    if (st) { parts.push({ f: st }); if (zip) parts.push(' '); }
    if (zip) parts.push({ f: zip });
    if (!parts.length) { MM.fieldDialog(); return; }
    insertParts(parts, 'Insert Address Block');
  };
  MM.fieldDialog = function () {
    if (!MM.data) return;
    const sel = h('select', { size: 10, id: 'mm-fields', style: 'width:100%' });
    for (const f of MM.data.fields) sel.appendChild(h('option', { value: f, text: f }));
    sel.selectedIndex = 0;
    sel.addEventListener('dblclick', () => { MM.insertField(sel.value); });
    ui.dialog({ title: 'Insert Merge Field', width: 300, body: h('div', { class: 'col' }, h('label', { for: 'mm-fields', text: 'Fields:' }), sel), buttons: [{ label: '&Insert', primary: true, onClick: () => { MM.insertField(sel.value); return false; } }, { label: 'Close' }] });
  };

  /* ---------- preview ---------- */
  function refreshFields() {
    if (!L.fields) return;
    const d = doc();
    const list = [];
    for (const st of D.stories(d)) D.walk(st, (b, c, i, story) => { if (b.t === 'p') for (const it of b.runs) if (it.t === 'fb' && /^\s*MERGEFIELD\b/i.test(it.instr)) list.push({ fid: it.fid, it, p: b, story }); });
    if (!list.length) return;
    const dirty = d.dirty;
    const n0 = D.hist.undo.length;
    D.tx('Preview Results', () => { for (const fr of list) L.fields.updateField(fr); });
    if (D.hist.undo.length > n0) D.hist.undo.pop();
    d.dirty = dirty;
  }
  MM.togglePreview = function () { MM.preview = !MM.preview; refreshFields(); refreshPane(); L.ui.refresh(); };
  MM.go = function (i) { const n = MM.records().length; if (!n) return; MM.rec = i < 0 ? n - 1 : L.clamp(i, 0, n - 1); if (!MM.preview) MM.preview = true; refreshFields(); refreshPane(); L.ui.refresh(); };
  MM.step = (dlt) => MM.go(MM.rec + dlt);

  /* ---------- merging ---------- */
  /** blocks for the current record with merge fields replaced by plain text */
  function mergedBlocks(srcBlocks) {
    const blocks = D.cloneBlocks(srcBlocks);
    D.walk({ blocks }, (b) => {
      if (b.t !== 'p') return;
      const out = [];
      let skip = null;
      for (const it of b.runs) {
        if (skip) { if (it.t === 'fe' && it.fid === skip) skip = null; continue; }
        if (it.t === 'fb' && /^\s*MERGEFIELD\b/i.test(it.instr)) {
          const f = L.fields.parse(it.instr);
          const v = MM.value(f.args[0] || '');
          const txt = (v && f.sw['\\b'] ? f.sw['\\b'] : '') + v + (v && f.sw['\\f'] ? f.sw['\\f'] : '');
          if (txt) out.push(D.text(L.fields.applyGeneral(txt, f.fmt.filter((x) => !/MERGEFORMAT/i.test(x))), L.clone(it.rPr || {})));
          skip = it.fid;
          continue;
        }
        if (it.t === 'fb' && /^\s*(NEXT|NEXTIF|SKIPIF|MERGEREC|MERGESEQ)\b/i.test(it.instr)) { skip = it.fid; if (/MERGEREC|MERGESEQ/i.test(it.instr)) out.push(D.text(String(MM.rec + 1), L.clone(it.rPr || {}))); continue; }
        out.push(it);
      }
      b.runs = out;
      D.normalize(b);
      /* address lines left empty by missing data are removed (Word's blank-line suppression) */
    });
    return blocks.filter((b) => !(b.t === 'p' && b._wasAddr && !D.plen(b)));
  }
  MM.buildMerged = function () {
    const src = doc();
    const recs = MM.records();
    if (!recs.length) return null;
    const out = D.newDoc();
    out.styles = L.clone(src.styles); out.numbering = L.clone(src.numbering); out.defaults = L.clone(src.defaults); out.settings = L.clone(src.settings); out.settings.track = false;
    out.hf = L.clone(src.hf); out.theme = L.clone(src.theme); out.sect = L.clone(src.sect); out.props = Object.assign({}, src.props, { title: (src.props.title || A().fileName) + ' (merged)' });
    out.main.blocks = [];
    const save = MM.rec;
    recs.forEach((ri, k) => {
      MM.rec = k;
      const bl = mergedBlocks(src.main.blocks);
      if (MM.type !== 'directory' && k < recs.length - 1) {
        /* each letter becomes its own section */
        const last = bl[bl.length - 1].t === 'p' ? bl[bl.length - 1] : (bl.push(D.para()), bl[bl.length - 1]);
        last.sect = Object.assign(L.clone(src.sect), { type: 'nextPage' });
      }
      out.main.blocks.push(...bl);
    });
    MM.rec = save;
    return out;
  };
  MM.mergeToNew = async function () {
    const d = MM.buildMerged();
    if (!d) { ui.msg('There are no recipients selected.', { icon: 'info' }); return; }
    A().untitled++;
    MM.preview = false;
    refreshFields();
    A().loadDoc(d, 'Letters' + A().untitled, { newWindow: true });
    MM.preview = false;
    A().status(`Merged ${MM.records().length} record(s) into a new document.`);
  };
  MM.mergeToPDF = async function () {
    const merged = MM.buildMerged();
    if (!merged) return;
    const keep = D.doc, keepHist = { u: D.hist.undo.slice(), r: D.hist.redo.slice() };
    try {
      D.doc = merged;
      LY.render();
      const blob = await A().buildPDF();
      L.ui.busy(false);
      D.doc = keep;
      D.hist.undo = keepHist.u; D.hist.redo = keepHist.r;
      LY.render();
      E.restoreDom(true);
      await L.saveFile(A().fileName + ' (merged).pdf', blob);
    } catch (e) { D.doc = keep; LY.render(); L.ui.busy(false); ui.msg('The merge could not be printed. ' + (e.message || ''), { icon: 'error' }); }
  };

  /* ---------- wizard (task pane) ---------- */
  function refreshPane() { if (L.panes && L.panes.task.current === 'mailmerge') L.panes.task.render(); }
  MM.renderWizard = function (b) {
    const { sec, link } = L.panes;
    const step = MM.wstep;
    const nav = h('div', { class: 'tp-btns', style: 'margin-top:14px' }, h('div', { class: 'tp-note', text: `Step ${step} of 6` }));
    if (step < 6) nav.appendChild(link('nextChange', 'Next: ' + ['Starting document', 'Select recipients', 'Write your letter', 'Preview your letters', 'Complete the merge'][step - 1], () => { MM.wstep++; if (MM.wstep === 5 && MM.data) MM.go(MM.rec); refreshPane(); }, { disabled: step === 3 && !MM.data }));
    if (step > 1) nav.appendChild(link('prevChange', 'Previous: ' + ['Select document type', 'Starting document', 'Select recipients', 'Write your letter', 'Preview your letters'][step - 2], () => { MM.wstep--; refreshPane(); }));
    if (step === 1) {
      const rad = (v, l) => ui.radio('mm-type', l, MM.type === v, () => { MM.type = v; });
      b.append(sec('Select document type', h('div', { class: 'tp-note', text: 'What type of document are you working on?' }), rad('letters', '&Letters'), rad('envelopes', '&Envelopes'), rad('labels', 'La&bels'), rad('directory', '&Directory')), h('div', { class: 'tp-note', text: 'Letters: send letters to a group of people. You can personalize the letter that each person receives.' }));
    } else if (step === 2) {
      b.append(sec('Select starting document', h('div', { class: 'tp-note', text: 'How do you want to set up your letters?' }), ui.radio('mm-start', 'Use the current &document', true, () => {}), ui.radio('mm-start', 'Start from a &template', false, () => L.dlg.newFromTemplate()), ui.radio('mm-start', 'Start from existing do&cument', false, () => ui.exec('open'))));
      if (MM.type === 'envelopes') b.appendChild(link('envelope', 'Envelope options...', () => L.dlg.envelopes()));
      if (MM.type === 'labels') b.appendChild(link('labels', 'Label options...', () => L.dlg.envelopes('labels')));
    } else if (step === 3) {
      b.append(sec('Select recipients', ui.radio('mm-src', 'Use an existing &list', !MM.data || !MM.data.typed, () => {}), ui.radio('mm-src', 'T&ype a new list', !!(MM.data && MM.data.typed), () => MM.typeList()),
        MM.data ? h('div', { class: 'tp-note', text: `Currently, your recipients are selected from: ${MM.data.name} (${MM.records().length} of ${MM.data.rows.length})` }) : h('div', { class: 'tp-note', text: 'Use names and addresses from a file or a database.' }),
        link('open', 'Browse...', () => MM.openSource()), MM.data ? link('mmRecipients', 'Edit recipient list...', () => MM.recipientsDialog()) : null, link('new', 'Type a new list...', () => MM.typeList())));
    } else if (step === 4) {
      b.append(sec('Write your letter', h('div', { class: 'tp-note', text: 'Write your letter now, then click a location in the document and choose one of the items below.' }), link('mmAddress', 'Address block...', () => MM.insertBlock('address')), link('mmGreeting', 'Greeting line...', () => MM.insertBlock('greeting')), link('mmField', 'More items...', () => MM.fieldDialog())));
    } else if (step === 5) {
      const n = MM.records().length;
      b.append(sec('Preview your letters', h('div', { class: 'tp-note', text: 'One of the merged letters is previewed here. To preview another letter, click one of the following:' }), h('div', { class: 'tp-row' }, ui.button('<<', () => MM.step(-1), { class: 'btn small' }), h('span', { text: `Recipient: ${MM.rec + 1} of ${n}` }), ui.button('>>', () => MM.step(1), { class: 'btn small' })), link('find', 'Find a recipient...', async () => { const q = await ui.prompt('Find:', '', 'Find Entry'); if (!q) return; const recs = MM.records(); const k = recs.findIndex((ri) => MM.data.rows[ri].some((v) => String(v).toLowerCase().includes(q.toLowerCase()))); if (k >= 0) MM.go(k); }), link('mmRecipients', 'Edit recipient list...', () => MM.recipientsDialog()), ui.button('Exclude this recipient', () => { const recs = MM.records(); MM.data.included.delete(recs[MM.rec]); MM.go(MM.rec); }, { class: 'btn small' })));
    } else {
      b.append(sec('Complete the merge', h('div', { class: 'tp-note', text: 'Quire is ready to produce your letters.' }), link('mmMergePrint', 'Print... (PDF)', () => MM.mergeToPDF()), link('mmMergeNew', 'Edit individual letters...', () => MM.mergeToNew())));
    }
    b.appendChild(nav);
  };
})();
