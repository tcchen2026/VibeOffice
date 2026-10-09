/* Quire — fields (evaluation, update, codes), bookmarks, footnotes & endnotes, captions,
 * cross-references, tables of contents / figures and indexes.
 */
(function () {
  'use strict';
  const L = window.L, D = L.D, O = L.O, E = L.ed, LY = L.layout;
  const F = (L.fields = {});
  const N = (L.notes = {});
  const doc = () => D.doc;
  const app = () => L.app;

  /* ================= instruction parsing ================= */
  /** split an instruction into tokens, honouring quotes: ['TOC', '\\o', '1-3', ...] */
  F.tokens = function (instr) {
    const out = [];
    const re = /\s*(?:"((?:[^"\\]|\\.)*)"|(\\[*#@!]|\\\w+)|([^\s"]+))/g;
    let m;
    const s = String(instr || '');
    while ((m = re.exec(s))) {
      if (m[0].trim() === '') break;
      if (m[1] != null) out.push({ q: true, v: m[1].replace(/\\"/g, '"').replace(/\\\\/g, '\\') });
      else if (m[2]) out.push({ sw: m[2] });
      else out.push({ v: m[3] });
    }
    return out;
  };
  /** {type, args:[...], sw:{'\\o':'1-3', '\\h':true, '\\*':['MERGEFORMAT']}} */
  F.parse = function (instr) {
    const toks = F.tokens(instr);
    const res = { type: '', args: [], sw: {}, fmt: [] };
    if (!toks.length) return res;
    const first = toks.shift();
    res.type = (first.v || '').toUpperCase();
    if (/^=/.test(first.v || '')) { res.type = '='; res.expr = String(instr).replace(/^\s*=/, '').replace(/\\[#*].*$/, '').trim(); }
    const argSw = ARGSW[res.type], optSw = OPTARG[res.type] || '';
    for (let i = 0; i < toks.length; i++) {
      const t = toks[i];
      if (t.sw) {
        const k = t.sw;
        const nx = toks[i + 1];
        const letter = k.slice(1).toLowerCase();
        if (k === '\\*') { res.fmt.push((nx && nx.v) || ''); i++; continue; }
        let takesArg;
        if (/^\\[#@]$/.test(k)) takesArg = !!nx;
        else if (!nx || nx.sw) takesArg = false;
        else if (argSw != null) takesArg = argSw.includes(letter) || (optSw.includes(letter) && !!nx.q);
        else takesArg = !!nx.q;
        if (takesArg) { res.sw[k] = nx.v; i++; } else res.sw[k] = true;
      } else if (res.type !== '=') res.args.push(t.v);
    }
    return res;
  };
  const ARGSW = { TOC: 'abcdflpst', INDEX: 'bcdefghklps', REF: 'd', SEQ: 'rs', XE: 'frty', TC: 'fl', MERGEFIELD: 'bf', HYPERLINK: 'lot', SYMBOL: 'fs', INCLUDEPICTURE: 'c', FILLIN: 'd', ASK: 'd', NOTEREF: '', PAGEREF: '', STYLEREF: '', DATE: '', TIME: '', PAGE: '', NUMPAGES: '', AUTHOR: '', TITLE: '', FILENAME: '' };
  const OPTARG = { TOC: 'no' };

  /* ================= formatting switches ================= */
  const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  /** Word date-time picture */
  F.formatDate = function (dt, pic) {
    if (!pic) pic = 'M/d/yyyy';
    pic = pic.replace(/^"|"$/g, '');
    let out = '';
    const re = /'([^']*)'|(dddd|ddd|dd|d|MMMM|MMM|MM|M|yyyy|yy|HH|H|hh|h|mm|m|ss|s|am\/pm|AM\/PM|a\/p|A\/P)|([\s\S])/g;
    let m;
    const H = dt.getHours();
    while ((m = re.exec(pic))) {
      if (m[1] != null) { out += m[1]; continue; }
      if (m[3] != null) { out += m[3]; continue; }
      switch (m[2]) {
        case 'dddd': out += DAYS[dt.getDay()]; break;
        case 'ddd': out += DAYS[dt.getDay()].slice(0, 3); break;
        case 'dd': out += String(dt.getDate()).padStart(2, '0'); break;
        case 'd': out += dt.getDate(); break;
        case 'MMMM': out += MONTHS[dt.getMonth()]; break;
        case 'MMM': out += MONTHS[dt.getMonth()].slice(0, 3); break;
        case 'MM': out += String(dt.getMonth() + 1).padStart(2, '0'); break;
        case 'M': out += dt.getMonth() + 1; break;
        case 'yyyy': out += dt.getFullYear(); break;
        case 'yy': out += String(dt.getFullYear()).slice(2); break;
        case 'HH': out += String(H).padStart(2, '0'); break;
        case 'H': out += H; break;
        case 'hh': out += String(H % 12 || 12).padStart(2, '0'); break;
        case 'h': out += H % 12 || 12; break;
        case 'mm': out += String(dt.getMinutes()).padStart(2, '0'); break;
        case 'm': out += dt.getMinutes(); break;
        case 'ss': out += String(dt.getSeconds()).padStart(2, '0'); break;
        case 's': out += dt.getSeconds(); break;
        case 'am/pm': out += H < 12 ? 'am' : 'pm'; break;
        case 'AM/PM': out += H < 12 ? 'AM' : 'PM'; break;
        case 'a/p': out += H < 12 ? 'a' : 'p'; break;
        case 'A/P': out += H < 12 ? 'A' : 'P'; break;
        default: out += m[2];
      }
    }
    return out;
  };
  F.DATE_PICTURES = ['M/d/yyyy', 'dddd, MMMM dd, yyyy', 'MMMM d, yyyy', 'M/d/yy', 'yyyy-MM-dd', 'd-MMM-yy', 'M.d.yyyy', 'MMM. d, yy', 'd MMMM yyyy', 'MMMM yy', 'MMM-yy', 'M/d/yyyy h:mm am/pm', 'M/d/yyyy h:mm:ss am/pm', 'h:mm am/pm', 'h:mm:ss am/pm', 'HH:mm', 'HH:mm:ss'];
  function applyGeneral(v, fmts) {
    let s = v;
    for (const f of fmts) {
      const k = (f || '').toUpperCase();
      const n = parseFloat(String(v).replace(/,/g, ''));
      if (k === 'UPPER') s = String(s).toUpperCase();
      else if (k === 'LOWER') s = String(s).toLowerCase();
      else if (k === 'FIRSTCAP') s = String(s).charAt(0).toUpperCase() + String(s).slice(1);
      else if (k === 'CAPS') s = L.titleCase ? L.titleCase(String(s)) : String(s).replace(/\b\w/g, (c) => c.toUpperCase());
      else if (!isNaN(n)) {
        if (f === 'ROMAN') s = D.fmtNum(n, 'upperRoman');
        else if (f === 'roman') s = D.fmtNum(n, 'lowerRoman');
        else if (f === 'ALPHABETIC') s = D.fmtNum(n, 'upperLetter');
        else if (f === 'alphabetic') s = D.fmtNum(n, 'lowerLetter');
        else if (k === 'ARABIC') s = String(n);
        else if (k === 'ORDINAL') s = D.fmtNum(n, 'ordinal');
        else if (k === 'CARDTEXT') s = D.fmtNum(n, 'cardinalText').toLowerCase();
        else if (k === 'ORDTEXT') s = D.fmtNum(n, 'ordinalText').toLowerCase();
        else if (k === 'HEX') s = Math.round(n).toString(16).toUpperCase();
        else if (k === 'DOLLARTEXT') { const w = Math.floor(n); s = D.fmtNum(w, 'cardinalText').toLowerCase() + ' and ' + String(Math.round((n - w) * 100)).padStart(2, '0') + '/100'; }
      }
    }
    return s;
  }
  F.applyGeneral = applyGeneral;

  /* ================= bookmarks ================= */
  /** {name → {p, o, endP, endO, story}} */
  F.bookmarks = function () {
    const d = doc();
    D.reindex(d);
    const out = new Map();
    const open = new Map();
    for (const st of D.stories(d)) for (const p of st.paras || []) {
      let o = 0;
      for (const it of p.runs) {
        if (it.t === 'bs' && it.name) { const b = { name: it.name, id: it.id, p, o, story: st }; out.set(it.name, b); open.set(it.id, b); }
        if (it.t === 'be' && open.has(it.id)) { const b = open.get(it.id); b.endP = p; b.endO = o; open.delete(it.id); }
        o += D.ilen(it);
      }
    }
    for (const b of out.values()) if (!b.endP) { b.endP = b.p; b.endO = b.o; }
    return out;
  };
  F.bookmarkText = function (name) {
    const b = F.bookmarks().get(name);
    if (!b) return null;
    return O.textRange(D.pos(b.p, b.o), D.pos(b.endP, b.endO), ' ');
  };
  F.addBookmark = function (name, a, b) {
    const d = doc();
    [a, b] = D.order(d, a, b);
    /* replace an existing bookmark of that name */
    const old = F.bookmarks().get(name);
    if (old) F.removeBookmarkItems(old);
    const id = String(D.nid());
    D.touch(a.p);
    const ia = D.splitAt(a.p, a.o);
    a.p.runs.splice(ia, 0, D.item('bs', { id, name }));
    D.touch(b.p);
    /* recompute index: the bs inserted before changes nothing for b offsets (zero length) */
    let ib = D.splitAt(b.p, b.o);
    if (a.p === b.p && ib <= ia) ib = ia + 1;
    b.p.runs.splice(ib, 0, D.item('be', { id }));
    return id;
  };
  F.removeBookmarkItems = function (bm) {
    for (const p of [bm.p, bm.endP]) { if (!p) continue; D.touch(p); p.runs = p.runs.filter((it) => !((it.t === 'bs' || it.t === 'be') && it.id === bm.id)); }
  };
  /** hidden bookmark around a paragraph's text (for TOC links and cross-references) */
  F.ensureParaBookmark = function (p, prefix) {
    for (const it of p.runs) if (it.t === 'bs' && it.name && it.name.startsWith(prefix)) return it.name;
    const name = prefix + String(100000000 + D.nid()).slice(1);
    let s = 0;
    const n = D.plen(p);
    F.addBookmark(name, D.pos(p, s), D.pos(p, n));
    return name;
  };

  /* ================= locating fields ================= */
  /** innermost field whose begin..end contains pos: {fid, it (fb), p, story} */
  F.at = function (pos) {
    if (!pos) return null;
    const d = doc();
    const st = D.storyOf(d, pos.p);
    const stack = [];
    let found = null;
    for (const p of st.paras || []) {
      let o = 0;
      for (const it of p.runs) {
        const before = p === pos.p && o >= pos.o;
        if (before && p === pos.p) { found = stack.length ? stack[stack.length - 1] : null; return found; }
        if (it.t === 'fb') stack.push({ fid: it.fid, it, p, story: st });
        else if (it.t === 'fe') { const i = stack.findIndex((x) => x.fid === it.fid); if (i >= 0) stack.splice(i, 1); }
        o += D.ilen(it);
      }
      if (p === pos.p) return stack.length ? stack[stack.length - 1] : null;
    }
    return null;
  };
  /** fields in the current selection (or the one around the caret), outermost first */
  F.inSelection = function () {
    if (!E.sel) return [];
    const d = doc();
    const st = E.story();
    if (E.collapsed()) { const f = F.at(E.sel.f); return f ? [f] : []; }
    const [a, b] = E.range();
    const out = [];
    for (const p of D.parasBetween(d, a.p, b.p)) {
      let o = 0;
      for (const it of p.runs) {
        const inside = (p !== a.p || o >= a.o) && (p !== b.p || o <= b.o);
        if (it.t === 'fb' && inside) out.push({ fid: it.fid, it, p, story: st });
        o += D.ilen(it);
      }
    }
    const around = F.at(a);
    if (around && !out.some((x) => x.fid === around.fid)) out.unshift(around);
    return out;
  };

  /* ================= evaluation ================= */
  const docProp = (k) => { const d = doc(); const v = d.props[k]; return v == null ? '' : String(v); };
  const user = () => L.store.get('userName', 'Quire User');
  /** compute a field result: string, array of blocks, or null (leave as is) */
  F.evaluate = function (instr, pos, opts) {
    opts = opts || {};
    const d = doc();
    const f = F.parse(instr);
    const t = f.type;
    const now = new Date();
    const fmt = (s) => applyGeneral(s, f.fmt.filter((x) => !/^(MERGEFORMAT|CHARFORMAT)$/i.test(x)));
    const num = (v) => { if (f.sw['\\#'] && L.tables) return L.tables.formatNumber(+v, f.sw['\\#']); return fmt(String(v)); };
    const date = (dt) => fmt(F.formatDate(dt, f.sw['\\@'] || (t === 'TIME' ? 'h:mm am/pm' : 'M/d/yyyy')));
    switch (t) {
      case 'DATE': case 'TIME': return date(now);
      case 'CREATEDATE': return date(new Date(d.props.created || now));
      case 'SAVEDATE': return date(new Date(d.props.modified || now));
      case 'PRINTDATE': return (d.props.printed || d.props.lastPrinted) ? date(new Date(d.props.printed || d.props.lastPrinted)) : date(new Date(0)).replace(/.*/, f.sw['\\@'] ? date(new Date(0)) : '0/0/0000');
      case 'EDITTIME': return num(Math.round((d.props.totalTime || 0)));
      case 'PAGE': case 'NUMPAGES': case 'SECTIONPAGES': case 'SECTION': {
        const pg = pos ? LY.pageOf(pos) : -1;
        const page = LY.pages[pg >= 0 ? pg : 0];
        return LY.dynField(t, f.fmt[0] || '', page) || '1';
      }
      case 'AUTHOR': return fmt(f.args[0] || docProp('creator'));
      case 'TITLE': return fmt(docProp('title'));
      case 'SUBJECT': return fmt(docProp('subject'));
      case 'KEYWORDS': return fmt(docProp('keywords'));
      case 'COMMENTS': return fmt(docProp('description'));
      case 'LASTSAVEDBY': return fmt(docProp('lastModifiedBy') || user());
      case 'REVNUM': return num(d.props.revision || 1);
      case 'TEMPLATE': return fmt(d.props.template || 'Normal.dot');
      case 'FILENAME': return fmt((app() && app().fileName ? app().fileName : 'Document1') + (f.sw['\\p'] ? '' : '') + '.docx');
      case 'FILESIZE': return num(0);
      case 'USERNAME': return fmt(f.args[0] || user());
      case 'USERINITIALS': return fmt(f.args[0] || L.store.get('userInitials', '') || O.initials());
      case 'USERADDRESS': return fmt(f.args[0] || L.store.get('userAddress', ''));
      case 'NUMWORDS': case 'NUMCHARS': case 'NUMPARAS': { const st = D.stats(D.allParas(d)); return num(t === 'NUMWORDS' ? st.words : t === 'NUMCHARS' ? st.charsNoSp : st.paras); }
      case 'DOCPROPERTY': { const k = (f.args[0] || '').toLowerCase(); const map = { title: 'title', subject: 'subject', author: 'creator', keywords: 'keywords', comments: 'description', company: 'company', manager: 'manager', category: 'category', lastsavedby: 'lastModifiedBy' }; if (map[k]) return fmt(docProp(map[k])); const c = d.custom && Object.keys(d.custom).find((x) => x.toLowerCase() === k); if (c) return fmt(String(d.custom[c].v != null ? d.custom[c].v : d.custom[c])); if (k === 'pages') return num(LY.pages.length || 1); if (k === 'words') return num(D.stats(D.allParas(d)).words); return 'Error! Unknown document property name.'; }
      case 'INFO': return F.evaluate((f.args[0] || 'TITLE') + ' ' + (f.args.slice(1).join(' ')), pos);
      case 'QUOTE': return fmt(f.args.join(' '));
      case 'SYMBOL': { const c = f.args[0] || ''; const code = /^0x/i.test(c) ? parseInt(c, 16) : parseInt(c, 10); return isNaN(code) ? '' : L.mapSymbolChar(String.fromCharCode(code), f.sw['\\f'] || 'Symbol') || String.fromCharCode(code); }
      case 'SEQ': return seqValue(f, pos, opts);
      case 'REF': case '': return refValue(f, pos);
      case 'PAGEREF': {
        const b = F.bookmarks().get(f.args[0]);
        if (!b) return 'Error! Bookmark not defined.';
        const pg = LY.pageOf(D.pos(b.p, b.o));
        if (pg < 0) return null;
        const page = LY.pages[pg];
        let s = D.fmtNum(page.numVal, page.numFmt);
        if (f.sw['\\p'] && pos) { const here = LY.pageOf(pos); s = here === pg ? 'above' : here > pg ? 'on page ' + s : 'on page ' + s; }
        return fmt(s);
      }
      case 'NOTEREF': {
        const b = F.bookmarks().get(f.args[0]);
        if (!b) return 'Error! Bookmark not defined.';
        let o = 0;
        for (const it of b.p.runs) { if (o >= b.o && (it.t === 'fn' || it.t === 'en')) return fmt(LY.noteNum ? LY.noteNum(it) : '1'); o += D.ilen(it); }
        return 'Error! Not a valid footnote reference.';
      }
      case 'STYLEREF': return styleRef(f, pos);
      case '=': {
        try {
          const v = L.tables.evalFormula(f.expr, pos ? pos.p : null);
          return L.tables.formatNumber(v, f.sw['\\#'] || null);
        } catch (e) { return '!' + (e.message || 'Syntax Error'); }
      }
      case 'IF': return ifValue(f, pos);
      case 'MERGEFIELD': {
        const name = f.args[0] || '';
        if (L.mailmerge && L.mailmerge.preview && L.mailmerge.data) { const v = L.mailmerge.value(name); return fmt((f.sw['\\b'] && v ? f.sw['\\b'] : '') + v + (f.sw['\\f'] && v ? f.sw['\\f'] : '')); }
        return '«' + name + '»';
      }
      case 'TOC': return opts.blocks ? F.buildTOC(f) : null;
      case 'INDEX': return opts.blocks ? F.buildIndex(f) : null;
      case 'LISTNUM': case 'AUTONUM': case 'AUTONUMLGL': case 'AUTONUMOUT': return null;
      case 'FILLIN': case 'ASK': return null;
      case 'HYPERLINK': case 'MACROBUTTON': case 'GOTOBUTTON': case 'INCLUDETEXT': case 'INCLUDEPICTURE': case 'LINK': case 'EMBED': case 'EQ': case 'ADVANCE': case 'XE': case 'TC': case 'RD': case 'TA': case 'PRIVATE': case 'ADDIN': case 'FORMTEXT': case 'FORMCHECKBOX': case 'FORMDROPDOWN': return null;
      default:
        /* a bare bookmark name behaves like REF */
        if (f.type && F.bookmarks().has(f.args.length ? f.type : instr.trim().split(/\s+/)[0])) return refValue({ args: [instr.trim().split(/\s+/)[0]], sw: f.sw, fmt: f.fmt }, pos);
        return 'Error! Bookmark not defined.';
    }
  };
  function refValue(f, pos) {
    const name = f.args[0];
    const b = F.bookmarks().get(name);
    if (!b) return 'Error! Reference source not found.';
    void pos;
    if (f.sw['\\n'] || f.sw['\\r'] || f.sw['\\w']) {
      const lb = LY.labels && LY.labels.get(b.p.id);
      if (lb) return applyGeneral(lb.text.replace(/[.)]\s*$/, ''), f.fmt);
      return '0';
    }
    let t = O.textRange(D.pos(b.p, b.o), D.pos(b.endP, b.endO), ' ');
    t = t.replace(/[\r\n\v]+/g, ' ');
    if (f.sw['\\p'] && pos) { const here = LY.pageOf(pos), there = LY.pageOf(D.pos(b.p, b.o)); t += here >= 0 && there >= 0 ? (here === there ? (pos.p._o < b.p._o ? ' below' : ' above') : ' on page ' + (there + 1)) : ''; }
    return applyGeneral(t, f.fmt.filter((x) => !/MERGEFORMAT/i.test(x)));
  }
  function seqValue(f, pos, opts) {
    const id = (f.args[0] || '').toUpperCase();
    const d = doc();
    if (!pos) return '1';
    let n = 0;
    const target = pos.p;
    let done = null;
    const stories = [d.main];
    for (const st of stories) {
      D.walk(st, (b) => {
        if (done != null || b.t !== 'p') return;
        /* heading-based reset \s level */
        if (f.sw['\\s']) { const lv = D.headingLevel(d, b); if (lv && lv <= +f.sw['\\s']) n = 0; }
        let o = 0;
        for (const it of b.runs) {
          if (it.t === 'fb') {
            const g = F.parse(it.instr);
            if (g.type === 'SEQ' && (g.args[0] || '').toUpperCase() === id) {
              if (g.sw['\\r'] != null) n = +g.sw['\\r'];
              else if (!g.sw['\\c']) n++;
              if (b === target && (opts.fid ? it.fid === opts.fid : o >= pos.o - 0)) { done = n; return false; }
            }
          }
          o += D.ilen(it);
        }
        return true;
      });
    }
    if (done == null) done = n || 1;
    if (f.sw['\\h']) return '';
    return applyGeneral(String(done), f.fmt.length ? f.fmt : []);
  }
  function styleRef(f, pos) {
    const d = doc();
    const name = (f.args[0] || '').toLowerCase();
    const st = D.findStyleByName(d, name) || d.styles[f.args[0]];
    const level = /^\d$/.test(name) ? +name : 0;
    const matches = (p) => (level ? D.headingLevel(d, p) === level : st && (p.pPr.style || 'Normal') === st.id);
    let p = pos && D.storyOf(d, pos.p) === d.main ? pos.p : null;
    if (!p) {
      /* in headers: first matching paragraph on the page being edited, else the nearest before */
      const pg = LY.hfEdit ? LY.pages[LY.hfEdit.page] : LY.pages[0];
      if (pg) { const frags = pg.body.querySelectorAll('.p'); for (const fr of frags) { const q = D.byId(d, +fr.dataset.pid); if (q && matches(q)) return D.plainText(q).trim(); } }
      p = d.main.paras && d.main.paras[0];
    }
    D.reindex(d);
    let q = p;
    while (q) { if (matches(q)) return D.plainText(q).trim(); q = D.prevPara(d, q); }
    return 'Error! No text of specified style in document.';
  }
  function ifValue(f, pos) {
    const toks = F.tokens(f.args.length ? '' : '');
    void toks;
    /* IF expr op expr "true" "false" (args already split) */
    const a = f.args;
    if (a.length < 3) return '';
    const val = (s) => { const n = parseFloat(s); return isNaN(n) ? s : n; };
    const ops = ['=', '<>', '<', '>', '<=', '>='];
    let i = ops.indexOf(a[1]) >= 0 ? 1 : -1;
    let left = a[0], op = a[1], right = a[2], rest = a.slice(3);
    if (i < 0) { const m = /^(.*?)(<>|<=|>=|=|<|>)(.*)$/.exec(a[0]); if (m) { left = m[1]; op = m[2]; right = m[3]; rest = a.slice(1); } }
    const L1 = val(left), R1 = val(right);
    let r;
    switch (op) { case '=': r = typeof R1 === 'string' && /[*?]/.test(R1) ? new RegExp('^' + R1.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.') + '$').test(String(L1)) : L1 == R1; break; case '<>': r = L1 != R1; break; case '<': r = L1 < R1; break; case '>': r = L1 > R1; break; case '<=': r = L1 <= R1; break; case '>=': r = L1 >= R1; break; default: r = false; } /* eslint-disable-line eqeqeq */
    void pos;
    return r ? rest[0] || '' : rest[1] || '';
  }

  /* ================= update ================= */
  F.updateField = function (fieldRef, opts) {
    const { it, story } = fieldRef;
    if (it.lock) return false;
    const f = F.parse(it.instr);
    if (it.ff) return false;
    const g = O.findField(story, it.fid);
    if (!g.begin) return false;
    const pos = D.pos(g.begin.p, g.begin.o);
    const res = F.evaluate(it.instr, pos, { blocks: true, fid: it.fid });
    if (res == null) return false;
    if (Array.isArray(res)) { O.setFieldResult(story, it.fid, res); return true; }
    /* keep the first result run's formatting (\* MERGEFORMAT behaviour) */
    O.setFieldResult(story, it.fid, String(res));
    void f; void opts;
    return true;
  };
  F.updateSelection = function () {
    let fields = F.inSelection();
    if (!fields.length) { app().status('There are no fields in the selection.'); return; }
    const hasToc = fields.some((x) => /^(TOC|INDEX)$/.test(F.parse(x.it.instr).type));
    E.edit('Update Field', () => {
      for (const fr of fields) F.updateField(fr);
      if (hasToc) F.refreshTOCPages(fields);
      return E.sel;
    });
  };
  F.updateAll = function (types) {
    const d = doc();
    const list = [];
    for (const st of D.stories(d)) D.walk(st, (b, c, i, story) => { if (b.t === 'p') for (const it of b.runs) if (it.t === 'fb' && (!types || types.includes(F.parse(it.instr).type))) list.push({ fid: it.fid, it, p: b, story }); });
    if (!list.length) return 0;
    E.edit('Update Fields', () => { for (const fr of list) F.updateField(fr); F.refreshTOCPages(list); return E.sel; });
    return list.length;
  };
  F.updateAllOfType = (type) => { const n = F.updateAll([type]); if (!n) app().status(type === 'TOC' ? 'There is no table of contents in this document.' : 'No fields were updated.'); };
  F.gotoFirst = function (type) {
    const d = doc();
    for (const st of [d.main]) for (const p of (D.reindex(d), st.paras)) { let o = 0; for (const it of p.runs) { if (it.t === 'fb' && F.parse(it.instr).type === type) { app().gotoPos(D.pos(p, o)); return; } o += D.ilen(it); } }
    app().status('There is no table of contents in this document.');
  };
  /** after rebuilding TOC/INDEX blocks: paginate once more and refresh PAGEREF results */
  F.refreshTOCPages = function (fields) {
    const tocs = fields.filter((x) => /^(TOC|INDEX)$/.test(F.parse(x.it.instr).type));
    if (!tocs.length) return;
    withPages(() => {
      D.reindex(doc());
      for (const st of D.stories(doc())) for (const p of st.paras || []) for (const it of p.runs) {
        if (it.t !== 'fb') continue;
        const t = F.parse(it.instr).type;
        if (t === 'PAGEREF') F.updateField({ fid: it.fid, it, p, story: st });
      }
      for (const t of tocs) if (F.parse(t.it.instr).type === 'INDEX') F.updateField(t);
    });
  };
  /** run fn with an up-to-date print layout (page numbers), restoring the current view */
  function withPages(fn) {
    const v = LY.view;
    if (v !== 'print') LY.view = 'print';
    LY.render();
    try { fn(); } finally { if (v !== 'print') { LY.view = v; } }
  }
  F.withPages = withPages;
  F.unlinkSelection = function () {
    const fields = F.inSelection();
    if (!fields.length) return;
    E.edit('Unlink Fields', () => {
      for (const fr of fields) {
        const g = O.findField(fr.story, fr.fid);
        for (const part of [g.begin, g.sep, g.end]) if (part) { D.touch(part.p); part.p.runs = part.p.runs.filter((x) => !((x.t === 'fb' || x.t === 'fs' || x.t === 'fe') && x.fid === fr.fid)); }
      }
      return E.sel;
    });
  };
  F.lockSelection = function (lock) {
    const fields = F.inSelection();
    if (!fields.length) return;
    E.edit(lock ? 'Lock Fields' : 'Unlock Fields', () => { for (const fr of fields) { D.touch(fr.p); fr.it.lock = lock || undefined; } return E.sel; });
  };
  /** replace a field's instruction (Edit Field) and update it */
  F.setInstr = function (fr, instr) {
    E.edit('Edit Field', () => { D.touch(fr.p); fr.it.instr = ' ' + instr.trim() + ' '; F.updateField(fr); return E.sel; });
  };
  /** fields Word updates before saving/printing (when the option is on) */
  F.beforeSave = function () {
    if (!(app() && app().opts.updateFieldsOnPrint)) return;
  };
  /* dynamic fields in headers refresh after layout */
  L.bus.on('layout-done', () => { if (LY.refreshDynamic) LY.refreshDynamic(); });

  /* ================= tables of contents, figures, index ================= */
  /** text width of the main section for right-aligned tab stops */
  const textWidth = () => { const s = doc().sect; return s.pgW - s.ml - s.mr - (s.gutter || 0); };
  const tocStyle = (lvl) => { const d = doc(); const id = 'TOC' + lvl; if (!d.styles[id]) { const b = D.builtinStyles()[id]; if (b) { D.touchKey(d, 'styles'); d.styles[id] = b; D.stylesChanged(); } } return id; };
  function pageText(p) {
    const pg = LY.pageOf(D.pos(p, 0));
    if (pg < 0) return '1';
    const page = LY.pages[pg];
    return D.fmtNum(page.numVal, page.numFmt);
  }
  /** entries for a TOC field: [{p, level, text}] */
  F.tocEntries = function (f) {
    const d = doc();
    const out = [];
    let lo = 1, hi = 9;
    const useOutline = f.sw['\\o'] != null;
    if (typeof f.sw['\\o'] === 'string') { const m = /(\d)\s*-\s*(\d)/.exec(f.sw['\\o']); if (m) { lo = +m[1]; hi = +m[2]; } }
    const styleMap = {};
    if (typeof f.sw['\\t'] === 'string') { const parts = f.sw['\\t'].split(/[,;]/); for (let i = 0; i + 1 < parts.length; i += 2) styleMap[parts[i].trim().toLowerCase()] = +parts[i + 1]; }
    const seqId = typeof f.sw['\\c'] === 'string' ? f.sw['\\c'].toUpperCase() : null;
    const tcId = typeof f.sw['\\f'] === 'string' ? f.sw['\\f'].toUpperCase() : f.sw['\\f'] ? '' : null;
    const lvlLimit = typeof f.sw['\\l'] === 'string' ? (/(\d)\s*-\s*(\d)/.exec(f.sw['\\l']) || []) : null;
    D.reindex(d);
    /* paragraphs belonging to existing TOC / INDEX results are not entries */
    const skip = new Set();
    const open = new Set();
    for (const b of d.main.paras) {
      let inside = open.size > 0;
      for (const it of b.runs) {
        if (it.t === 'fb' && /^(TOC|INDEX)$/.test(F.parse(it.instr).type)) { open.add(it.fid); inside = true; }
        else if (it.t === 'fe' && open.has(it.fid)) open.delete(it.fid);
      }
      if (inside) skip.add(b);
    }
    for (const b of d.main.paras) {
      if (skip.has(b)) continue;
      const text = tocText(b);
      if (seqId) {
        if (b.runs.some((it) => it.t === 'fb' && F.parse(it.instr).type === 'SEQ' && (F.parse(it.instr).args[0] || '').toUpperCase() === seqId)) out.push({ p: b, level: 1, text: f.sw['\\a'] ? stripLabel(text) : text });
        continue;
      }
      if (tcId != null) for (const it of b.runs) if (it.t === 'fb' && F.parse(it.instr).type === 'TC') {
        const g = F.parse(it.instr);
        if ((g.sw['\\f'] || '').toUpperCase() === tcId || (!g.sw['\\f'] && tcId === '')) { const lv = +(g.sw['\\l'] || 1); if (!lvlLimit || (lv >= +lvlLimit[1] && lv <= +lvlLimit[2])) out.push({ p: b, level: lv, text: g.args.join(' ') }); }
      }
      if (!text.trim()) continue;
      const sname = (D.styleDisplayName(d.styles[b.pPr.style || 'Normal']) || '').toLowerCase();
      if (styleMap[sname]) { out.push({ p: b, level: styleMap[sname], text }); continue; }
      if (useOutline || f.sw['\\u']) {
        const lv = f.sw['\\u'] && b.pPr.outline != null ? b.pPr.outline + 1 : D.headingLevel(d, b);
        if (lv && lv >= lo && lv <= hi) out.push({ p: b, level: lv, text });
      }
    }
    return out;
  };
  const tocFids = new Set();
  function tocText(p) { return D.plainText(p, { skipDeleted: true }).replace(/[\t\n]+/g, ' ').replace(/\s+$/, ''); }
  function stripLabel(t) { return t.replace(/^\S+\s+[\dA-Za-z.-]+[:.]?\s*/, ''); }
  F.buildTOC = function (f) {
    const d = doc();
    for (const p of D.allParas(d)) for (const it of p.runs) if (it.t === 'fb' && /^(TOC|INDEX)$/.test(F.parse(it.instr).type)) tocFids.add(it.fid);
    const entries = F.tocEntries(f);
    const links = f.sw['\\h'];
    const tw = textWidth();
    const noPg = typeof f.sw['\\n'] === 'string' ? (/(\d)\s*-\s*(\d)/.exec(f.sw['\\n']) || []) : f.sw['\\n'] ? [0, 1, 9] : null;
    const sepTab = f.sw['\\p'] == null;
    const isFig = typeof f.sw['\\c'] === 'string';
    const blocks = [];
    if (!entries.length) {
      blocks.push(D.para([D.text(isFig ? 'No table of figures entries found.' : 'No table of contents entries found.', { b: true })]));
      return blocks;
    }
    for (const e of entries) {
      const bm = F.ensureParaBookmark(e.p, '_Toc');
      const style = isFig ? (d.styles.TableofFigures ? 'TableofFigures' : (tocStyle(1))) : tocStyle(e.level);
      if (isFig && !d.styles.TableofFigures) { D.touchKey(d, 'styles'); d.styles.TableofFigures = D.builtinStyles().TableofFigures; D.stylesChanged(); }
      const rp = links ? { style: 'Hyperlink', link: { anchor: bm } } : {};
      if (links && !d.styles.Hyperlink) { D.touchKey(d, 'styles'); d.styles.Hyperlink = D.builtinStyles().Hyperlink; }
      /* TOC hyperlinks keep the TOC style's look (Word uses \h without visible link formatting) */
      if (links) delete rp.style;
      const runs = [D.text(e.text, rp)];
      const showPg = !(noPg && e.level >= +noPg[1] && e.level <= +noPg[2]);
      if (showPg) {
        runs.push(sepTab ? D.item('tab', null, links ? { link: { anchor: bm } } : {}) : D.text(String(f.sw['\\p']), rp));
        const fid = D.newFid();
        const base = links ? { link: { anchor: bm } } : {};
        runs.push(D.item('fb', { fid, instr: ` PAGEREF ${bm} \\h ` }, L.clone(base)), D.item('fs', { fid }, L.clone(base)), D.text(pageText(e.p), L.clone(base)), D.item('fe', { fid }, L.clone(base)));
      }
      blocks.push(D.para(runs, { style, tabs: sepTab && showPg ? [{ pos: L.round(tw, 2), al: 'right', leader: 'dot' }] : undefined }));
    }
    return blocks;
  };
  /** XE entries → index paragraphs */
  F.buildIndex = function (f) {
    const d = doc();
    const entries = new Map(); /* main → {pages:Set, subs: Map} */
    const sepPages = typeof f.sw['\\l'] === 'string' ? f.sw['\\l'] : ', ';
    const sepEntry = typeof f.sw['\\e'] === 'string' ? f.sw['\\e'] : ', ';
    const headingFmt = typeof f.sw['\\h'] === 'string' ? f.sw['\\h'] : null;
    const runIn = !!f.sw['\\r'];
    D.reindex(d);
    for (const p of d.main.paras) {
      let o = 0;
      for (const it of p.runs) {
        if (it.t === 'fb') {
          const g = F.parse(it.instr);
          if (g.type === 'XE') {
            const raw = g.args.join(' ');
            const parts = raw.split(':');
            const pg = LY.pageOf(D.pos(p, o));
            const pageStr = g.sw['\\t'] ? String(g.sw['\\t']) : pg >= 0 ? D.fmtNum(LY.pages[pg].numVal, LY.pages[pg].numFmt) : '1';
            const fmtPg = { s: pageStr, b: !!g.sw['\\b'], i: !!g.sw['\\i'], see: !!g.sw['\\t'] };
            const main = parts[0].trim();
            if (!entries.has(main)) entries.set(main, { pages: [], subs: new Map() });
            const ent = entries.get(main);
            if (parts[1]) { const sub = parts.slice(1).join(':').trim(); if (!ent.subs.has(sub)) ent.subs.set(sub, []); const list = ent.subs.get(sub); if (!list.some((x) => x.s === fmtPg.s)) list.push(fmtPg); }
            else if (!ent.pages.some((x) => x.s === fmtPg.s)) ent.pages.push(fmtPg);
          }
        }
        o += D.ilen(it);
      }
    }
    const blocks = [];
    if (!entries.size) { blocks.push(D.para([D.text('No index entries found.', { b: true })])); return blocks; }
    const ensure = (id) => { if (!d.styles[id]) { D.touchKey(d, 'styles'); d.styles[id] = D.builtinStyles()[id]; D.stylesChanged(); } return id; };
    const pageRuns = (list) => { const out = []; list.forEach((pg, i) => { if (i) out.push(D.text(sepPages)); out.push(D.text(pg.see ? pg.s : pg.s, { b: pg.b || undefined, i: pg.i || pg.see || undefined })); }); return out; };
    const keys = Array.from(entries.keys()).sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' }));
    let lastLetter = null;
    for (const k of keys) {
      const letter = k.charAt(0).toUpperCase();
      if (headingFmt && letter !== lastLetter) {
        lastLetter = letter;
        const h = headingFmt.replace(/A/g, letter);
        if (blocks.length) blocks.push(D.para([], { style: ensure('IndexHeading') }));
        blocks.push(D.para([D.text(h)], { style: ensure('IndexHeading') }));
      }
      const ent = entries.get(k);
      const runs = [D.text(k)];
      if (ent.pages.length) runs.push(D.text(sepEntry), ...pageRuns(ent.pages));
      if (runIn && ent.subs.size) {
        runs.push(D.text(': '));
        Array.from(ent.subs.keys()).sort().forEach((s, i) => { if (i) runs.push(D.text('; ')); runs.push(D.text(s + sepEntry), ...pageRuns(ent.subs.get(s))); });
        blocks.push(D.para(runs, { style: ensure('Index1') }));
        continue;
      }
      blocks.push(D.para(runs, { style: ensure('Index1') }));
      for (const s of Array.from(ent.subs.keys()).sort((a, b) => a.localeCompare(b))) blocks.push(D.para([D.text(s + sepEntry), ...pageRuns(ent.subs.get(s))], { style: ensure('Index2') }));
    }
    return blocks;
  };
  /** insert a TOC / TOF / INDEX field with its result at the caret */
  F.insertGenerated = function (instr, opts) {
    opts = opts || {};
    const f = F.parse(instr);
    let inserted;
    E.edit(f.type === 'INDEX' ? 'Insert Index' : 'Insert Table of Contents', () => {
      let pos = E.deleteSelection();
      /* fields producing paragraphs start at the beginning of a paragraph */
      if (pos.o > 0) pos = O.splitPara(pos);
      if (D.plen(pos.p) > 0) { O.splitPara(pos); }
      const fid = D.newFid();
      D.touch(pos.p);
      const p = pos.p;
      p.runs.unshift(D.item('fb', { fid, instr: ' ' + instr.trim() + ' ' }), D.item('fs', { fid }), D.item('fe', { fid }));
      tocFids.add(fid);
      const story = D.storyOf(doc(), p);
      inserted = { it: p.runs[0], story };
      D.reindex(doc());
      const blocks = f.type === 'INDEX' ? F.buildIndex(f) : F.buildTOC(f);
      O.setFieldResult(story, fid, blocks);
      if (opts.columns > 1 && f.type === 'INDEX') {
        /* Word surrounds a multi-column index with continuous section breaks */
        const d = doc();
        D.reindex(d);
        const g = O.findField(story, fid);
        const tb0 = D.topBlock(d, g.begin.p), tb1 = D.topBlock(d, g.end.p);
        const cont = D.touchList(d.main);
        const before = D.para([], {});
        const curSect = D.sectOfBlock(d, tb0.i).sect;
        // The original section's start belongs before the index. Both the index
        // and the following remainder start continuously within that section.
        before.sect = L.opc.duplicate(curSect);
        const endP = cont.blocks[tb1.i];
        D.touch(endP);
        endP.sect = Object.assign(endP.sect || L.opc.duplicate(curSect), { type: 'continuous', cols: { n: opts.columns, space: 36, eq: true, w: [] } });
        D.touchKey(curSect, 'type'); curSect.type = 'continuous';
        cont.blocks.splice(tb0.i, 0, before);
        d._idxDirty = true;
      }
      return E.sel;
    });
    if (inserted) E.edit('Update Page Numbers', () => { F.refreshTOCPages([inserted]); return E.sel; }, { merge: null });
  };

  /* ================= captions & cross-references ================= */
  F.CAPTION_LABELS = ['Equation', 'Figure', 'Table'];
  F.insertCaption = function (opts) {
    const d = doc();
    const label = opts.label || 'Figure';
    const fmt = opts.fmt && opts.fmt !== 'ARABIC' ? ' \\* ' + opts.fmt : ' \\* ARABIC';
    const instr = `SEQ ${label.replace(/\s+/g, '_')}${fmt}`;
    E.edit('Insert Caption', () => {
      if (!d.styles.Caption) { D.touchKey(d, 'styles'); d.styles.Caption = D.builtinStyles().Caption; D.stylesChanged(); }
      const pos = E.collapsed() ? E.sel.f : E.range()[0];
      let p;
      const ci = D.cellOf(d, pos.p);
      const tbl = ci ? ci.tbl : null;
      /* caption paragraph above or below the selected item / table / current paragraph */
      if (tbl && !opts.inPlace) {
        const tb = D.tblInfo(d, tbl);
        const cont = D.touchList(tb.cont);
        const i = cont.blocks.indexOf(tbl);
        p = D.para([], { style: 'Caption' });
        cont.blocks.splice(opts.position === 'above' ? i : i + 1, 0, p);
      } else {
        const cont = D.touchList(D.info(d, pos.p).cont);
        const i = cont.blocks.indexOf(pos.p);
        if (opts.inPlace) p = pos.p;
        else { p = D.para([], { style: 'Caption' }); cont.blocks.splice(opts.position === 'above' ? i : i + 1, 0, p); }
      }
      d._idxDirty = true;
      D.reindex(d);
      D.touch(p);
      let at = D.pos(p, D.plen(p));
      at = O.insertText(at, label + ' ', {});
      if (opts.chapter) {
        at = O.insertField(at, `STYLEREF ${opts.chapterStyle || 1} \\s`, '1', {});
        at = O.insertText(at, opts.chapterSep || '-', {});
      }
      at = O.insertField(at, instr, '1', {});
      if (opts.text) at = O.insertText(at, opts.text, {});
      return at;
    });
    /* number it */
    const fields = [];
    D.reindex(d);
    for (const p of D.allParas(d)) for (const it of p.runs) if (it.t === 'fb' && /^(SEQ|STYLEREF)$/.test(F.parse(it.instr).type)) fields.push({ fid: it.fid, it, p, story: d.main });
    E.edit('Update Captions', () => { for (const fr of fields) F.updateField(fr); return E.sel; });
  };
  /** cross-reference targets of a kind */
  F.xrefItems = function (kind) {
    const d = doc();
    D.reindex(d);
    const out = [];
    if (kind === 'Heading') { for (const p of d.main.paras) { const lv = D.headingLevel(d, p); if (lv) out.push({ p, text: ' '.repeat((lv - 1) * 2) + D.plainText(p).trim(), level: lv }); } }
    else if (kind === 'Bookmark') { for (const [name, b] of F.bookmarks()) if (!name.startsWith('_')) out.push({ name, text: name, b }); }
    else if (kind === 'Numbered item') { for (const p of d.main.paras) { const lb = LY.labels && LY.labels.get(p.id); if (lb && lb.lv.fmt !== 'bullet') out.push({ p, text: lb.text + ' ' + D.plainText(p).trim() }); } }
    else if (kind === 'Footnote' || kind === 'Endnote') {
      const k = kind === 'Footnote' ? 'fn' : 'en';
      for (const p of d.main.paras) { let o = 0; for (const it of p.runs) { if (it.t === k && !it.self) { const st = (k === 'fn' ? d.fn : d.en)[it.id]; out.push({ p, o, it, text: (LY.noteNum ? LY.noteNum(it) : '') + ' ' + (st ? D.plainText(D.firstPara(st.blocks) || D.para()).trim() : '') }); } o += D.ilen(it); } }
    } else {
      /* caption labels */
      for (const p of d.main.paras) for (const it of p.runs) if (it.t === 'fb') { const g = F.parse(it.instr); if (g.type === 'SEQ' && (g.args[0] || '').replace(/_/g, ' ').toLowerCase() === kind.toLowerCase()) { out.push({ p, text: D.plainText(p).trim(), seq: true }); break; } }
    }
    return out;
  };
  /** insert a cross-reference field. what: 'text'|'page'|'number'|'abovebelow'|'labelnum'|'caption'|'notenum' */
  F.insertXref = function (kind, item, what, opts) {
    opts = opts || {};
    const d = doc();
    E.edit('Insert Cross-reference', () => {
      let name;
      if (kind === 'Bookmark') name = item.name;
      else if (kind === 'Footnote' || kind === 'Endnote') {
        name = '_Ref' + String(100000000 + D.nid()).slice(1);
        F.addBookmark(name, D.pos(item.p, item.o), D.pos(item.p, item.o + 1));
      } else if (item.seq && (what === 'labelnum')) {
        /* bookmark the label and number only */
        let end = 0, o = 0, seen = false;
        for (const it of item.p.runs) { o += D.ilen(it); if (it.t === 'fb' && F.parse(it.instr).type === 'SEQ') seen = true; if (seen && it.t === 'fe') { end = o; break; } }
        name = '_Ref' + String(100000000 + D.nid()).slice(1);
        F.addBookmark(name, D.pos(item.p, 0), D.pos(item.p, end || D.plen(item.p)));
      } else name = F.ensureParaBookmark(item.p, '_Ref');
      const sw = opts.link !== false ? ' \\h' : '';
      let instr;
      switch (what) {
        case 'page': instr = `PAGEREF ${name}${sw}`; break;
        case 'number': instr = `REF ${name} \\r${sw}`; break;
        case 'numberFull': instr = `REF ${name} \\w${sw}`; break;
        case 'abovebelow': instr = `REF ${name} \\p${sw}`; break;
        case 'notenum': instr = `NOTEREF ${name}${sw}`; break;
        default: instr = `REF ${name}${sw}`;
      }
      if (opts.abovebelow && what !== 'abovebelow') instr += ' \\p';
      const pos = E.deleteSelection();
      const res = F.evaluate(instr, pos);
      return O.insertField(pos, instr, res == null ? '' : String(res), O.inheritRPr(pos));
    });
    void d;
  };

  /* ================= footnotes & endnotes ================= */
  N.insert = function (kind, opts) {
    opts = opts || {};
    const d = doc();
    if (!E.sel || D.storyOf(d, E.sel.f.p) !== d.main) { app().status('Notes can only be inserted in the main document.'); return; }
    let target = null;
    E.edit(kind === 'fn' ? 'Insert Footnote' : 'Insert Endnote', () => {
      const store = kind === 'fn' ? 'fn' : 'en';
      D.touchKey(d, store);
      let id = 1;
      while (d[store][id] != null) id++;
      id = String(id);
      const refStyle = kind === 'fn' ? 'FootnoteReference' : 'EndnoteReference';
      const textStyle = kind === 'fn' ? 'FootnoteText' : 'EndnoteText';
      if (!d.styles[refStyle] || !d.styles[textStyle]) { D.touchKey(d, 'styles'); const b = D.builtinStyles(); d.styles[refStyle] = d.styles[refStyle] || b[refStyle]; d.styles[textStyle] = d.styles[textStyle] || b[textStyle]; D.stylesChanged(); }
      const custom = opts.mark || null;
      const self = D.item(kind, { id, self: true }, { style: refStyle });
      if (custom) self.custom = custom;
      const np = D.para([self, D.text(' ')], { style: textStyle });
      d[store][id] = { kind: store, id, blocks: [np] };
      const ref = D.item(kind, { id }, Object.assign(O.inheritRPr(E.sel.f), { style: refStyle }));
      if (custom) ref.custom = custom;
      const pos = E.deleteSelection();
      O.insertItem(pos, ref);
      target = np;
      d._idxDirty = true;
      return E.sel;
    });
    if (target) {
      D.reindex(d);
      setTimeout(() => { E.setSel(D.pos(target, D.plen(target))); E.focus(); }, 0);
    }
  };
  /** notes settings: {fmt, start, restart, pos} for footnotes or endnotes */
  N.setOptions = function (kind, o) {
    const d = doc();
    E.edit('Note Options', () => { D.touchKey(d, 'settings'); const k = kind === 'fn' ? 'fnPr' : 'enPr'; d.settings[k] = Object.assign({}, d.settings[k], o); return E.sel; });
  };
  /** convert all footnotes to endnotes or the reverse */
  N.convertAll = function (from) {
    const d = doc();
    const to = from === 'fn' ? 'en' : 'fn';
    E.edit('Convert Notes', () => {
      D.touchKey(d, 'fn'); D.touchKey(d, 'en');
      const map = {};
      let next = 1;
      for (const id of Object.keys(d[from])) {
        while (d[to][next] != null) next++;
        const nid = String(next++);
        const st = d[from][id];
        st.kind = to; st.id = nid;
        D.walk(st, (b) => { if (b.t === 'p') for (const it of b.runs) if (it.t === from && it.self) { it.t = to; it.id = nid; it.rPr = Object.assign({}, it.rPr, { style: to === 'fn' ? 'FootnoteReference' : 'EndnoteReference' }); } if (b.t === 'p' && b.pPr.style === (from === 'fn' ? 'FootnoteText' : 'EndnoteText')) b.pPr.style = to === 'fn' ? 'FootnoteText' : 'EndnoteText'; });
        d[to][nid] = st;
        map[id] = nid;
      }
      d[from] = {};
      D.walk(d.main, (b) => { if (b.t === 'p') { let changed = false; for (const it of b.runs) if (it.t === from && !it.self && map[it.id]) { if (!changed) { D.touch(b); changed = true; } it.t = to; it.id = map[it.id]; it.rPr = Object.assign({}, it.rPr, { style: to === 'fn' ? 'FootnoteReference' : 'EndnoteReference' }); } } });
      d._idxDirty = true;
      return E.sel;
    });
  };
  /** remove orphaned notes (reference deleted) — runs before saving */
  N.prune = function () {
    const d = doc();
    const used = { fn: new Set(), en: new Set() };
    D.walk(d.main, (b) => { if (b.t === 'p') for (const it of b.runs) if ((it.t === 'fn' || it.t === 'en') && !it.self) used[it.t].add(it.id); });
    for (const k of ['fn', 'en']) for (const id of Object.keys(d[k])) if (!used[k].has(id)) delete d[k][id];
  };
})();
