// The English UI text and its translations (docs/i18n.md).
//
//   node tools/i18n.mjs gather        scan the code for the English the UI shows and write the catalog,
//                                     public/common/i18n/en.json ({ "<English>": { "where": [...] } }); report the
//                                     text that can't be translated yet (English pieces joined together, a value
//                                     inside a template literal) in tmp/i18n/untranslatable.txt
//   node tools/i18n.mjs status        for each language: translated, missing and obsolete entries, placeholders that
//                                     differ; tmp/i18n/<lang>-todo.json lists the missing English for the translator
//   node tools/i18n.mjs chunks <lang> [size]   the missing English of a language in pieces for the translator,
//                                     tmp/i18n/<lang>/todo-NN.json: [["<English>", "<where>"], ...], in the order of
//                                     where the text is, so a dialog's strings come together
//   node tools/i18n.mjs merge <lang> <files>   adds the translations in files ({ "<English>": "<translation>" }) to
//                                     public/common/i18n/<lang>.json; an entry not in the catalog, or whose
//                                     placeholders or access key mark differ from the English's, is left out and listed
//   node tools/i18n.mjs prune <lang>  removes the entries the catalog no longer has from public/common/i18n/<lang>.json
//   node tools/i18n.mjs harvest [apps] [--lang xx]   opens each app (and the Start Center) in the pseudo-locale, then
//                                     every menu and submenu, task pane, ribbon tab and dialog (the commands named "…"),
//                                     and lists the English still on screen, outside the document (tmp/i18n/harvest-
//                                     <app>.txt), and the text its box cuts off (cut-<app>-<lang>.txt); with --lang, a
//                                     language's translation, only what is cut off
//
// gather is a static scan: a string literal is UI text when it sits where the UI takes text (a label, tip or
// title property, an argument of L.t, L.N (text in a table, translated where shown) or of the ui.* helpers, a child of h(), the label of a ui.select option, an
// AutoShape name), or when it reads as an English sentence or label (several words, or ending in "...", ":" or
// "?", beginning with a capital); plus the pages' own HTML text and Ledger's function descriptions. A string the
// scan misses shows in English in the pseudo-locale (?lang=qps), which is how to find it.
import fs from 'node:fs';
import path from 'node:path';
import { pathToFileURL } from 'node:url';

const root = path.resolve(path.dirname(new URL(import.meta.url).pathname), '..');
const PUB = path.join(root, 'public');
const LANGS = ['ja', 'ko', 'fr', 'de', 'zh-CN', 'zh-Hant'];
const catalogFile = path.join(PUB, 'common/i18n/en.json');

/* files with no UI text of their own: readers, writers, engines (text they put in files or errors they throw are
   handled where they are shown); and the licences page's script, whose page is in English only, as the licences are */
const SKIP = new Set(['oss-licenses.js', 'common/crypto.js', 'common/sha.js', 'common/zip.js', 'common/xml.js', 'common/opc.js', 'common/opc-order.js',
  'common/numfmt.js', 'common/dml.js', 'common/metafile.js', 'common/pdf.js', 'common/icons.js', 'common/i18n.js', 'sw.js',
  'quire/js/docx-read.js', 'quire/js/docx-write.js', 'quire/js/rtf.js', 'quire/js/htmlio.js', 'quire/js/markdown.js', 'quire/js/mdio.js',
  'quire/js/omml.js', 'quire/js/preserve.js', 'quire/js/preserve-drawing.js', 'quire/js/dmodel.js',
  'ledger/js/formula.js', 'ledger/js/calc.js', 'ledger/js/fn-core.js', 'ledger/js/fn-lookup.js', 'ledger/js/fn-stat.js', 'ledger/js/fn-fin.js',
  'ledger/js/fn-eng.js', 'ledger/js/xlsx-read.js', 'ledger/js/xlsx-write.js', 'ledger/js/xml.js', 'ledger/js/xmlss.js', 'ledger/js/extensions.js', 'ledger/js/preserve.js',
  'lectern/js/pptx-read.js', 'lectern/js/pptx-write.js', 'lectern/js/preserve.js', 'lectern/js/properties.js', 'lectern/js/smartart-io.js']);
const HTML = ['index.html', 'lite.html', 'quire/index.html', 'lectern/index.html'];   // Ledger's page is built from its template
const LEDGER_TEMPLATE = 'tools/ledger/build/template.html';

/* where the UI takes text */
const PROPS = new Set(['label', 'menuLabel', 'tbLabel', 'tip', 'title', 'text', 'placeholder', 'heading', 'note', 'aria-label', 'data-tip', 'what', 'drop', 'caption', 'hint']);
const CALLS = { t: [0], 'L.t': [0], 'VO.t': [0], N: [0], 'L.N': [0], 'VO.N': [0], 'ui.msg': [0], 'ui.toast': [0], 'ui.busy': [1], 'ui.prompt': [0, 2], 'ui.button': [0], 'ui.check': [0],
  'ui.radio': [1], 'ui.field': [0], 'ui.group': [0], 'ui.tbDrop': [0, 3], 'ui.toolbar': [1], 'G.f': [0], 'def': [1, 3], 'O.tx': [1] };
const H = new Set(['h', 'L.h']);
/* properties whose text is not shown (the assistant's descriptions of his own moves, ids, styles) */
const HIDDEN = new Set(['desc', 'id', 'cmd', 'icon', 'key', 'class', 'style', 'type', 'font', 'family', 'tag', 'xml', 'rel', 'ct']);
/* the templates' own content stays English this round (docs/i18n.md): only their names and categories */
const TEMPLATES = { 'quire/js/templates.js': null, 'lectern/js/templates.js': null, 'ledger/js/panes.js': /^  const TPL = \{$[\s\S]*?^  \};$/m };
/* content put into documents (AutoContent's slides, the sample deck and workbook, the rows of a scenario summary,
   a PivotTable or subtotals): English this round, like the templates' content, so only what is marked (L.N, L.t)
   or passed to a call that shows it (an undo step's name) counts */
const CONTENT = {
  'lectern/js/app.js': [/^  const AUTOCONTENT = \{$[\s\S]*?^  \};$/m, /^  function buildSample\(\) \{$[\s\S]*?^  \}$/m],
  'ledger/js/app.js': [/^  A\.sampleWorkbook = function \(\) \{$[\s\S]*?^  \};$/m],
  'ledger/js/dialogs2.js': [/^    const summary = \(\) => \{$[\s\S]*?^    \};$/m, /^  D\.buildPivot = function [\s\S]*?^  \};$/m, /^ *s\.put\(gt, st\.at, .*$/m],
};

/* ---------------------------------------------------------------- a small JavaScript tokenizer */
const KW = new Set(['return', 'typeof', 'instanceof', 'in', 'of', 'new', 'delete', 'void', 'throw', 'case', 'do', 'else', 'yield', 'await']);
const tokenize = (src) => lex(src, 0, 1, false).out;
/** the tokens from src[i0] on; with untilBrace, up to the } that closes a template literal's ${ (returned as i) */
function lex(src, i0, line0, untilBrace) {
  const out = [];
  let i = i0, line = line0, depth = 0;
  const n = src.length;
  const prev = () => out[out.length - 1];
  const regexOK = () => { const p = prev(); return !p || (p.type === 'p' && !')]}'.includes(p.value)) || (p.type === 'id' && KW.has(p.value)); };
  while (i < n) {
    const c = src[i];
    if (c === '\n') { line++; i++; continue; }
    if (c === ' ' || c === '\t' || c === '\r') { i++; continue; }
    if (c === '/' && src[i + 1] === '/') { while (i < n && src[i] !== '\n') i++; continue; }
    if (c === '/' && src[i + 1] === '*') { const e = src.indexOf('*/', i + 2), end = e < 0 ? n : e + 2; for (let k = i; k < end; k++) if (src[k] === '\n') line++; i = end; continue; }
    if (c === '"' || c === "'") {
      let j = i + 1, v = '';
      while (j < n && src[j] !== c) {
        if (src[j] === '\\') { const e = src[j + 1]; v += e === 'n' ? '\n' : e === 't' ? '\t' : e === 'u' ? String.fromCharCode(parseInt(src.substr(j + 2, 4), 16)) : e; j += e === 'u' ? 6 : 2; continue; }
        if (src[j] === '\n') break;
        v += src[j++];
      }
      out.push({ type: 'str', value: v, line });
      i = j + 1;
      continue;
    }
    if (c === '`') {
      /* a template literal: its text with each ${…} as {n}; the expressions are tokenized too */
      let j = i + 1, v = '', k = 0, startLine = line;
      const inner = [];
      while (j < n && src[j] !== '`') {
        if (src[j] === '\\') { v += src[j + 1] === 'n' ? '\n' : src[j + 1]; j += 2; continue; }
        if (src[j] === '$' && src[j + 1] === '{') {
          const sub = lex(src, j + 2, line, true);   // the expression, with its own strings, regexes and templates
          inner.push(...sub.out);
          v += '{' + k++ + '}';
          j = sub.i + 1; line = sub.line;
          continue;
        }
        if (src[j] === '\n') line++;
        v += src[j++];
      }
      out.push({ type: k ? 'tpl' : 'str', value: v, line: startLine });
      out.push(...inner.map((tk) => ({ ...tk, inTpl: true })));
      i = j + 1;
      continue;
    }
    if (c === '/' && regexOK()) {
      let j = i + 1, cls = false;
      while (j < n && (cls || src[j] !== '/')) { if (src[j] === '\\') j++; else if (src[j] === '[') cls = true; else if (src[j] === ']') cls = false; else if (src[j] === '\n') break; j++; }
      j++;
      while (j < n && /[a-z]/.test(src[j])) j++;
      out.push({ type: 're', line });
      i = j;
      continue;
    }
    const id = /^[A-Za-z_$][\w$]*/.exec(src.slice(i, i + 64));
    if (id) { out.push({ type: 'id', value: id[0], line }); i += id[0].length; continue; }
    const num = /^\d[\w.]*/.exec(src.slice(i, i + 32));
    if (num) { out.push({ type: 'num', line }); i += num[0].length; continue; }
    if (untilBrace && c === '{') depth++;
    if (untilBrace && c === '}' && !depth--) break;
    out.push({ type: 'p', value: c, line });
    i++;
  }
  return { out, i, line };
}

/* ---------------------------------------------------------------- which strings are UI text */
/* "Line|chart": UI text with its context (common/i18n.js) */
const context = /(?<=\S)\|[A-Za-z][A-Za-z -]*$/;   // as common/i18n.js plain()
const looksLikeCode = (s) => /[<>{};=\\|]|:\/\/|^[.#[@$%^]|^\s|\s$|^[\w-]+:[\w-]|["']\s*,|\.(js|css|html|xml|json|png|svg|jpg|docx|xlsx|pptx)\b|^[a-z]+([A-Z][a-z]*)+$/.test(s.replace(/\{\d+\}/g, '').replace(context, ''));
const hasWords = (s) => /[A-Za-z]{2}/.test(s.replace(/&(?=[A-Za-z])/g, ''));   // "B&y:" has words
/* "Format Cells", "Save changes?", "Size:", "&Undo", "Insert Picture..." — not "Times New Roman"-like single names */
const sentence = (t) => { const s = t.replace(context, ''); return /^[&"(“]?[A-Z]/.test(s) && /[a-z]/.test(s) && (/ /.test(s) || /(\.\.\.|…|:|\?|!)$/.test(s) || /&[A-Za-z]/.test(s)); };

function scanJS(file, src, add, fragment, markedOnly) {
  const toks = tokenize(src);
  /* the lines of template content, where only names and categories are UI text */
  let tpl = null;
  if (file in TEMPLATES) {
    const m = TEMPLATES[file] ? TEMPLATES[file].exec(src) : { index: 0, 0: src };
    if (m) { const a = src.slice(0, m.index).split('\n').length; tpl = [a, a + m[0].split('\n').length - 1]; }
  }
  const content = (CONTENT[file] || []).map((re) => re.exec(src)).filter(Boolean).map((m) => { const a = src.slice(0, m.index).split('\n').length; return [a, a + m[0].split('\n').length - 1]; });
  const stack = [];   // { kind: '(' | '{' | '[', callee, arg, key, afterColon, size }
  let chain = [];      // the dotted name just read: ui . msg
  for (let i = 0; i < toks.length; i++) {
    const tk = toks[i];
    const top = stack[stack.length - 1];
    if (tk.type === 'id') { const p = toks[i - 1]; chain = p && p.type === 'p' && p.value === '.' && chain.length ? [...chain, tk.value] : [tk.value]; continue; }
    if (tk.type === 'p') {
      const v = tk.value;
      if (v === '(') stack.push({ kind: '(', callee: toks[i - 1] && toks[i - 1].type === 'id' ? chain.join('.') : null, arg: 0 });
      else if (v === '{') stack.push({ kind: '{', key: null });
      else if (v === '[') stack.push({ kind: '[', size: 0 });
      else if (v === ')' || v === '}' || v === ']') stack.pop();
      else if (v === ',' && top) { if (top.kind === '(') top.arg++; else if (top.kind === '{') top.key = null; else if (top.kind === '[') top.size++; }
      /* a property name follows { or , (the ':' of a ternary inside an object is not one) */
      else if (v === ':' && top && top.kind === '{') { const k = toks[i - 1], b = toks[i - 2]; if (k && (k.type === 'id' || k.type === 'str') && b && b.type === 'p' && (b.value === '{' || b.value === ',')) top.key = k.value; }
      if (v !== '.') chain = [];
      continue;
    }
    if (tk.type !== 'str' && tk.type !== 'tpl') continue;
    const s = tk.value, next = toks[i + 1], prev = toks[i - 1];
    if (next && next.type === 'p' && next.value === ':' && top && top.kind === '{' && prev && prev.type === 'p' && (prev.value === '{' || prev.value === ',')) continue;   // a property name
    if (!hasWords(s)) continue;
    if (top && top.kind === '{' && HIDDEN.has(top.key)) continue;
    if ((prev && prev.type === 'p' && prev.value === '=' && toks[i - 2] && toks[i - 2].type === 'p' && '=!'.includes(toks[i - 2].value)) || (next && next.type === 'p' && '=!'.includes(next.value) && toks[i + 2] && toks[i + 2].value === '=')) continue;   // compared: x === 'formats'
    if (tpl && tk.line >= tpl[0] && tk.line <= tpl[1]) {
      /* not the sheets' names of Ledger's templates, which go into the workbook */
      if (top && top.kind === '{' && (top.key === 'name' || top.key === 'cat') && !stack.some((x) => x.kind === '{' && x.key === 'sheets')) add(s, file + ':' + tk.line);
      continue;
    }
    /* where it sits */
    let ui = false;
    if (top && top.kind === '{' && PROPS.has(top.key)) ui = true;
    if (top && top.kind === '(' && top.callee) {
      const name = top.callee.replace(/^(window|root)\./, '').replace(/^L\.ui\./, 'ui.');
      if (Object.hasOwn(CALLS, name) && CALLS[name].includes(top.arg)) ui = true;   // own keys: a class's constructor() is no call of Object's
      if (H.has(name) && top.arg >= 2) ui = true;
    }
    /* [['id', 'Label'], …] in the first argument of ui.select */
    const outer = stack[stack.length - 3];
    if (top && top.kind === '[' && top.size === 1 && stack[stack.length - 2] && stack[stack.length - 2].kind === '[' && outer && outer.kind === '(' && /(^|\.)select$/.test(outer.callee || '') && outer.arg === 0) ui = true;
    /* an explicit L.t / L.N argument is UI text whatever it holds; elsewhere markup, CSS and code are not */
    const explicit = top && top.kind === '(' && /^(t|N|L\.t|L\.N|VO\.t|VO\.N)$/.test(top.callee || '');
    if (markedOnly && !explicit) continue;
    if (!explicit && !(ui && top.kind === '(') && content.some(([a, b]) => tk.line >= a && tk.line <= b)) continue;
    if (!ui && !sentence(s)) continue;
    /* markup: a tag with attributes or a closing tag ('<blank>' is a label) */
    const markup = /<[a-z][\w-]*\s+[\w-]+=|<\/[a-z]/i.test(s);
    if (explicit ? markup : ui ? markup || /^[a-z-]+:[^;]*;|:\/\/|^[.#]\w/i.test(s) : looksLikeCode(s)) continue;
    const joined = (prev && prev.type === 'p' && prev.value === '+') || (next && next.type === 'p' && next.value === '+');
    const where = file + ':' + tk.line;
    if (tk.type === 'tpl' || joined) fragment(s, where);
    else add(s, where);
  }
}

/* Ledger's function descriptions: NAME|arguments|description lines (ledger/js/fninfo.js) */
function scanFninfo(file, src, add) {
  for (const m of src.matchAll(/^\s*[A-Z][A-Z0-9.]*\|[^|\n]*\|([^\n`]+)$/gm)) add(m[1].trim(), file + ':' + (src.slice(0, m.index).split('\n').length));
}

/* the pages' own text: text between tags and the attributes the UI shows (common/i18n.js translateTree) */
function scanHTML(file, src, add) {
  /* not inside translate="no" (an element with none of its own name inside: the Start Center's list of language names) */
  const body = src.replace(/<script[\s\S]*?<\/script>/g, (m) => m.replace(/[^\n]/g, ' ')).replace(/<style[\s\S]*?<\/style>/g, (m) => m.replace(/[^\n]/g, ' '))
    .replace(/<(\w+)\b[^>]*\stranslate="no"[^>]*>[\s\S]*?<\/\1>/g, (m) => m.replace(/[^\n]/g, ' '));
  const lineOf = (i) => body.slice(0, i).split('\n').length;
  for (const m of body.matchAll(/>([^<]+)</g)) {
    const s = m[1].replace(/\s+/g, ' ').trim().replace(/&amp;/g, '&').replace(/&nbsp;/g, '\u00a0');
    if (hasWords(s) && !/^[\w.-]+$/.test(s) || /^[A-Z][a-z]+$/.test(s)) add(s, file + ':' + lineOf(m.index));
  }
  for (const m of body.matchAll(/\s(?:title|placeholder|aria-label|data-tip|alt)="([^"]+)"/g)) if (hasWords(m[1])) add(m[1].replace(/&amp;/g, '&'), file + ':' + lineOf(m.index));
}

function gather() {
  const cat = new Map(), frags = [];
  const add = (s, where) => { if (!cat.has(s)) cat.set(s, []); const w = cat.get(s); if (w.length < 4 && !w.includes(where)) w.push(where); };
  const fragment = (s, where) => frags.push(where + '\t' + JSON.stringify(s));
  const walk = (dir) => fs.readdirSync(dir, { withFileTypes: true }).flatMap((e) => (e.isDirectory() ? walk(path.join(dir, e.name)) : [path.join(dir, e.name)]));
  for (const abs of walk(PUB).filter((f) => f.endsWith('.js')).sort()) {
    const rel = path.relative(PUB, abs);
    if (rel.startsWith('common/dict/')) continue;
    const src = fs.readFileSync(abs, 'utf8');
    if (SKIP.has(rel) || rel.startsWith('common/legacy/')) { scanJS(rel, src, add, fragment, true); continue; }   // only what is marked: L.t, L.N, N (the Office 97-2003 importers are readers too)
    if (rel === 'ledger/js/fninfo.js') scanFninfo(rel, src, add);
    scanJS(rel, src, add, fragment);
  }
  for (const rel of HTML) {
    const src = fs.readFileSync(path.join(PUB, rel), 'utf8');
    scanHTML(rel, src, add);   // no page has a script of its own (tools/csp.mjs): the Start Center's is start.js
  }
  scanHTML('ledger/index.html', fs.readFileSync(path.join(root, LEDGER_TEMPLATE), 'utf8'), add);
  const out = {};
  for (const k of [...cat.keys()].sort()) out[k] = { where: cat.get(k) };
  return { catalog: out, fragments: frags };
}


/* ---------------------------------------------------------------- harvest: English left on screen in the pseudo-locale */
/* the document itself, which keeps its own text */
const DOC_AREAS = { quire: '#scroller, #docmap, #revpane', ledger: '#gridwrap, #tabs', lectern: '#leftpane, #editorwrap, #notes', start: '#recentlist, .recent, .tiles' };
/* commands not to run: they print, leave the page or ask the computer for a file */
const DONT_RUN = /print|^(exit|quit|close|closeAll)$|^open|insertFile|fromFile|^save$|send|mail$/i;
const COLLECT = `(scope, skip, where) => {
  const out = window.__harvest = window.__harvest || new Map();
  const english = (s) => /[A-Za-z]{2}/.test(s) && !/\\[/.test(s);
  const note = (s, kind) => { s = s.replace(/\\s+/g, ' ').trim(); if (!s || (kind !== 'cut' && !english(s))) return; const k = kind + '\\t' + s; if (!out.has(k)) out.set(k, where); };
  /* text its box cuts off (or ends with "…"): a control too narrow for the language */
  const cut = (el) => { const cs = getComputedStyle(el); return (cs.overflowX !== 'visible' || cs.textOverflow === 'ellipsis') && el.scrollWidth > el.clientWidth + 1; };
  for (const top of document.querySelectorAll(scope)) {
    if (!top.getClientRects().length) continue;
    const w = document.createTreeWalker(top, NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT);
    for (let n = w.currentNode; n; n = w.nextNode()) {
      const el = n.nodeType === 1 ? n : n.parentElement;
      if (!el || (skip && el.closest(skip)) || el.closest('[contenteditable="true"], [translate="no"], script, style, svg, .tooltip')) continue;
      /* a label split by its access key (<u>) is judged whole */
      if (n.nodeType === 3) { const whole = el.closest('u') || (n.nextSibling && n.nextSibling.nodeName === 'U') || (n.previousSibling && n.previousSibling.nodeName === 'U') ? el.closest('u') ? el.parentElement : el : el; if (el.getClientRects().length && english(whole === el ? n.data : whole.textContent)) note(whole.textContent, 'text'); if (el.getClientRects().length && n.data.trim() && cut(whole)) note(whole.textContent, 'cut'); }
      else for (const a of ['title', 'aria-label', 'data-tip', 'placeholder']) if (n.getAttribute(a)) note(n.getAttribute(a), a);
    }
  }
  return out.size;
}`;
async function harvest(args) {
  const i = args.indexOf('--lang'), lang = i >= 0 ? args[i + 1] : 'qps', apps = args.filter((a, k) => i < 0 || (k !== i && k !== i + 1));
  const { ensureServer, openPage } = await import('./ooxml/cdp.mjs');
  const stop = await ensureServer();
  const pause = (ms) => new Promise((r) => setTimeout(r, ms));
  for (const app of apps.length ? apps : ['quire', 'ledger', 'lectern', 'start']) {
    const page = await openPage(app === 'start' ? '' : app, { persistence: false, automated: true, lang });
    const ev = (js) => page.evaluate(js);
    try {
      await page.send('Page.setInterceptFileChooserDialog', { enabled: true }).catch(() => {});
      await page.send('Emulation.setDeviceMetricsOverride', { width: 1280, height: 900, deviceScaleFactor: 1, mobile: false });
      /* no native boxes that would wait for an answer */
      await ev('window.alert = () => {}; window.confirm = () => false; window.prompt = () => null; true');
      await pause(500);
      const skip = JSON.stringify(DOC_AREAS[app] || '');
      const collect = (scope, where) => ev(`(${COLLECT})(${JSON.stringify(scope)}, ${skip}, ${JSON.stringify(where)})`);
      await collect('body', 'the window');
      if (app !== 'start') {
        /* menus and their submenus */
        const menus = await ev('(L.ui.menus || []).length');
        for (let i = 0; i < menus; i++) {
          await ev(`document.querySelectorAll('.mb-item')[${i}].dispatchEvent(new PointerEvent('pointerdown', { bubbles: true, cancelable: true }))`); await pause(150);
          const name = await ev(`L.ui.plainAmp(L.ui.menus[${i}].label)`);
          await collect('.menu', 'menu ' + name);
          const subs = await ev(`document.querySelectorAll('.menu .m-item.has-sub').length`);
          for (let j = 0; j < subs; j++) {
            await ev(`(() => { const r = document.querySelectorAll('.menu')[0].querySelectorAll('.m-item.has-sub')[${j}]; if (r) r.dispatchEvent(new PointerEvent('pointerenter')); })()`); await pause(150);
            await collect('.menu', 'menu ' + name + ' › submenu');
          }
          await ev('L.ui.closeMenus()'); await pause(80);
        }
        /* task panes */
        for (const id of await ev('Object.keys(L.panes.task.registry)')) {
          await ev(`(() => { try { L.panes.task.show(${JSON.stringify(id)}); } catch (e) {} })()`); await pause(250);
          await collect('#taskpane', 'task pane ' + id);
        }
        /* the ribbon (Paper 2016's layout, for this page only) */
        await ev(`document.documentElement.dataset.look = 'paper2016'; document.documentElement.dataset.layout = 'ribbon'; window.dispatchEvent(new Event('vo-look'))`); await pause(300);
        const tabs = await ev(`document.querySelectorAll('.rb-tab').length`);
        for (let i = 0; i < tabs; i++) {
          const id = await ev(`(() => { const t = document.querySelectorAll('.rb-tab')[${i}]; t.click(); return t.dataset.tab; })()`); await pause(200);
          await collect('.ribbon, .rb-backstage, .backstage', 'ribbon tab ' + id);
          await ev(`document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }))`); await pause(80);
        }
        await ev(`document.documentElement.dataset.look = 'classic'; document.documentElement.dataset.layout = 'bars'; window.dispatchEvent(new Event('vo-look'))`); await pause(200);
        /* the dialogs: each command named "…" */
        const ids = await ev(`Object.values(L.ui.cmds).filter((c) => c.label && /(\\.\\.\\.|…)$/.test(L.ui.plainAmp(c.label)) && (!c.enabled || c.enabled())).map((c) => c.id)`);
        for (const id of ids.filter((x) => !DONT_RUN.test(x))) {
          await ev(`(() => { try { void L.ui.exec(${JSON.stringify(id)}); } catch (e) {} })()`); await pause(350);
          if (await ev(`document.querySelectorAll('.dlg').length`)) {
            await collect('.dlg', 'dialog ' + id);
            const tabsN = await ev(`document.querySelectorAll('.dlg .tab').length`);
            for (let k = 0; k < tabsN; k++) { await ev(`(() => { const t = document.querySelectorAll('.dlg .tab')[${k}]; if (t) t.click(); })()`); await pause(120); await collect('.dlg', 'dialog ' + id + ' › tab ' + (k + 1)); }
          }
          await collect('#taskpane', 'after ' + id);
          for (let n = 0; n < 6 && await ev(`document.querySelectorAll('.dlg').length`); n++) { await ev(`(() => { const d = [...document.querySelectorAll('.dlg')].pop(); const x = d && d.querySelector('.dlg-x:not(.dlg-help)'); if (x) x.click(); })()`); await pause(150); }
          await ev('L.ui.closeMenus()');
        }
      } else {
        for (const id of ['settings', 'about']) { await ev(`(() => { const b = document.getElementById(${JSON.stringify(id)}); if (b) b.click(); })()`); await pause(300); await collect('dialog[open]', 'dialog ' + id); await ev(`document.querySelectorAll('dialog[open]').forEach((d) => d.close())`); }
      }
      const all = await ev('[...(window.__harvest || new Map())].map(([k, w]) => w + "\\t" + k)');
      const cuts = all.filter((x) => x.split('\t')[1] === 'cut'), found = all.filter((x) => x.split('\t')[1] !== 'cut');
      write(path.join(root, 'tmp/i18n', `cut-${app}-${lang}.txt`), cuts.sort().join('\n') + '\n');
      if (lang !== 'qps') { console.log(`${app} (${lang}): ${cuts.length} texts cut off (tmp/i18n/cut-${app}-${lang}.txt)`); continue; }
      /* what the code asked VO.t for that the catalog lacks */
      const missed = await ev('[...VO.i18n.missed]');
      write(path.join(root, 'tmp/i18n', 'missed-' + app + '.txt'), missed.sort().join('\n') + '\n');
      write(path.join(root, 'tmp/i18n', 'harvest-' + app + '.txt'), found.sort().join('\n') + '\n');
      console.log(`${app}: ${found.length} English texts on screen (tmp/i18n/harvest-${app}.txt), ${missed.length} asked for and not in the catalog (missed-${app}.txt), ${cuts.length} cut off (cut-${app}-qps.txt)${page.errors.length ? '; page errors: ' + page.errors.length : ''}`);
    } finally { await page.close(); }
  }
  stop();
}

const write = (file, text) => { fs.mkdirSync(path.dirname(file), { recursive: true }); fs.writeFileSync(file, text); };
const cmd = import.meta.url === pathToFileURL(process.argv[1] || '').href ? process.argv[2] : null;   // run, not imported by the test
if (cmd === null) { /* imported */ } else if (cmd === 'gather') {
  const { catalog, fragments } = gather();
  write(catalogFile, JSON.stringify(catalog, null, 1) + '\n');
  write(path.join(root, 'tmp/i18n/untranslatable.txt'), fragments.join('\n') + '\n');
  console.log(`${path.relative(root, catalogFile)}: ${Object.keys(catalog).length} strings; ${fragments.length} pieces of text that can't be translated yet (tmp/i18n/untranslatable.txt)`);
} else if (cmd === 'status') {
  const catalog = JSON.parse(fs.readFileSync(catalogFile, 'utf8'));
  const marks = (s) => (s.match(/\{\d+\}/g) || []).sort().join();
  for (const lang of LANGS) {
    const file = path.join(PUB, 'common/i18n', lang + '.json');
    const dict = fs.existsSync(file) ? JSON.parse(fs.readFileSync(file, 'utf8')) : {};
    const missing = Object.keys(catalog).filter((k) => typeof dict[k] !== 'string' || !dict[k]);
    const obsolete = Object.keys(dict).filter((k) => !(k in catalog));
    const placeholders = Object.keys(dict).filter((k) => k in catalog && typeof dict[k] === 'string' && dict[k] && marks(dict[k]) !== marks(k));
    write(path.join(root, 'tmp/i18n', lang + '-todo.json'), JSON.stringify(Object.fromEntries(missing.map((k) => [k, catalog[k]])), null, 1) + '\n');
    console.log(`${lang}: ${Object.keys(catalog).length - missing.length} of ${Object.keys(catalog).length} translated, ${missing.length} missing (tmp/i18n/${lang}-todo.json), ${obsolete.length} obsolete, ${placeholders.length} with other placeholders`);
  }
} else if (cmd === 'chunks' && LANGS.includes(process.argv[3])) {
  const lang = process.argv[3], size = +process.argv[4] || 300;
  const catalog = JSON.parse(fs.readFileSync(catalogFile, 'utf8'));
  const file = path.join(PUB, 'common/i18n', lang + '.json'), dict = fs.existsSync(file) ? JSON.parse(fs.readFileSync(file, 'utf8')) : {};
  const order = (w) => w.replace(/:(\d+)$/, (m, n) => ':' + n.padStart(6, '0'));
  const todo = Object.keys(catalog).filter((k) => !dict[k]).map((k) => [k, catalog[k].where[0]]).sort((a, b) => (order(a[1]) < order(b[1]) ? -1 : 1));
  const dir = path.join(root, 'tmp/i18n', lang);
  fs.rmSync(dir, { recursive: true, force: true });
  for (let i = 0; i * size < todo.length; i++) write(path.join(dir, `todo-${String(i + 1).padStart(2, '0')}.json`), JSON.stringify(todo.slice(i * size, (i + 1) * size), null, 0).replace(/\],\[/g, '],\n[') + '\n');
  console.log(`${lang}: ${todo.length} strings to translate in ${Math.ceil(todo.length / size)} files, tmp/i18n/${lang}/todo-NN.json`);
} else if (cmd === 'merge' && LANGS.includes(process.argv[3])) {
  const lang = process.argv[3];
  const catalog = JSON.parse(fs.readFileSync(catalogFile, 'utf8'));
  const file = path.join(PUB, 'common/i18n', lang + '.json'), dict = fs.existsSync(file) ? JSON.parse(fs.readFileSync(file, 'utf8')) : {};
  const marks = (s) => (s.match(/\{\d+\}/g) || []).sort().join();
  const keys = (s) => (s.replace(/&&/g, '').match(/&[^\s&]/g) || []).length;
  const problems = [];
  let added = 0;
  for (const f of process.argv.slice(4)) {
    for (const [en, tr] of Object.entries(JSON.parse(fs.readFileSync(f, 'utf8')))) {
      const why = !(en in catalog) ? 'not in the catalog' : typeof tr !== 'string' || !tr.trim() ? 'empty' : marks(tr) !== marks(en) ? 'placeholders differ' : keys(en) !== keys(tr) ? 'access key mark differs' : null;
      if (why) { problems.push(`${path.basename(f)}\t${why}\t${JSON.stringify(en)}\t${JSON.stringify(tr)}`); continue; }
      if (dict[en] !== tr) added++;
      dict[en] = tr;
    }
  }
  write(file, JSON.stringify(Object.fromEntries(Object.keys(dict).sort().map((k) => [k, dict[k]])), null, 1) + '\n');
  write(path.join(root, 'tmp/i18n', lang + '-problems.txt'), problems.join('\n') + '\n');
  console.log(`${lang}: ${added} translations added or changed, ${Object.keys(dict).length} in all; ${problems.length} left out (tmp/i18n/${lang}-problems.txt)`);
} else if (cmd === 'prune' && LANGS.includes(process.argv[3])) {
  const lang = process.argv[3], file = path.join(PUB, 'common/i18n', lang + '.json');
  const catalog = JSON.parse(fs.readFileSync(catalogFile, 'utf8')), dict = JSON.parse(fs.readFileSync(file, 'utf8'));
  const gone = Object.keys(dict).filter((k) => !(k in catalog));
  for (const k of gone) delete dict[k];
  write(file, JSON.stringify(dict, null, 1) + '\n');
  console.log(`${lang}: ${gone.length} entries the catalog no longer has removed` + (gone.length ? ': ' + gone.map((k) => JSON.stringify(k)).join(', ') : ''));
} else if (cmd === 'harvest') {
  await harvest(process.argv.slice(3));
} else {
  console.log('usage: node tools/i18n.mjs gather | status | chunks <lang> [size] | merge <lang> <files> | prune <lang> | harvest [apps] [--lang xx]');
  process.exit(1);
}

export { gather, tokenize };
