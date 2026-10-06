/* Ledger — formula language.
 * Tokeniser and Pratt parser for Excel A1 formulas → AST; serialiser; reference arithmetic used by
 * copy/fill, insert/delete rows and columns, moves and sheet renames.
 * References keep absolute positions (0-based) plus $ flags, so a formula can be re-printed for any cell.
 */
(function (root) {
  'use strict';
  const L = root.L || (root.L = {});
  const F = (L.formula = {});
  const MAXR = (F.MAXR = 1048576), MAXC = (F.MAXC = 16384);

  /* ---------------------------------------------------------------- A1 */
  F.colName = function (c) { let s = ''; c = c + 1; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); } return s; };
  F.colIndex = function (s) { let n = 0; s = s.toUpperCase(); for (let i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64); return n - 1; };
  F.cellName = (r, c, ra, ca) => (ca ? '$' : '') + F.colName(c) + (ra ? '$' : '') + (r + 1);
  /** "B3" → {r:2,c:1}; "$B$3" too; null when invalid */
  F.parseCell = function (s) {
    const m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})$/.exec(String(s).trim());
    if (!m) return null;
    const c = F.colIndex(m[2]), r = +m[4] - 1;
    if (c < 0 || c >= MAXC || r < 0 || r >= MAXR) return null;
    return { r, c, ca: !!m[1], ra: !!m[3] };
  };
  /** "A1:B5", "A:A", "3:3", "B2" → {r1,c1,r2,c2} (normalised) or null */
  F.parseRange = function (s) {
    s = String(s).trim().replace(/\$/g, '');
    let m;
    if ((m = /^([A-Za-z]{1,3}\d{1,7})(?::([A-Za-z]{1,3}\d{1,7}))?$/.exec(s))) {
      const a = F.parseCell(m[1]), b = m[2] ? F.parseCell(m[2]) : a;
      if (!a || !b) return null;
      return { r1: Math.min(a.r, b.r), c1: Math.min(a.c, b.c), r2: Math.max(a.r, b.r), c2: Math.max(a.c, b.c) };
    }
    if ((m = /^([A-Za-z]{1,3}):([A-Za-z]{1,3})$/.exec(s))) { const a = F.colIndex(m[1]), b = F.colIndex(m[2]); if (a >= MAXC || b >= MAXC) return null; return { r1: 0, c1: Math.min(a, b), r2: MAXR - 1, c2: Math.max(a, b) }; }
    if ((m = /^(\d{1,7}):(\d{1,7})$/.exec(s))) { const a = +m[1] - 1, b = +m[2] - 1; if (a < 0 || b < 0 || a >= MAXR || b >= MAXR) return null; return { r1: Math.min(a, b), c1: 0, r2: Math.max(a, b), c2: MAXC - 1 }; }
    return null;
  };
  F.rangeName = function (rg) {
    if (rg.r1 === 0 && rg.r2 === MAXR - 1 && !(rg.c1 === 0 && rg.c2 === MAXC - 1)) return F.colName(rg.c1) + ':' + F.colName(rg.c2);
    if (rg.c1 === 0 && rg.c2 === MAXC - 1) return (rg.r1 + 1) + ':' + (rg.r2 + 1);
    const a = F.colName(rg.c1) + (rg.r1 + 1);
    return rg.r1 === rg.r2 && rg.c1 === rg.c2 ? a : a + ':' + F.colName(rg.c2) + (rg.r2 + 1);
  };
  F.absRangeName = (rg) => {
    if (rg.r1 === 0 && rg.r2 === MAXR - 1) return '$' + F.colName(rg.c1) + ':$' + F.colName(rg.c2);
    if (rg.c1 === 0 && rg.c2 === MAXC - 1) return '$' + (rg.r1 + 1) + ':$' + (rg.r2 + 1);
    const a = '$' + F.colName(rg.c1) + '$' + (rg.r1 + 1);
    return rg.r1 === rg.r2 && rg.c1 === rg.c2 ? a : a + ':$' + F.colName(rg.c2) + '$' + (rg.r2 + 1);
  };
  /** sheet name as it must appear in a formula */
  F.quoteSheet = function (name) {
    if (name == null) return '';
    if (/^[A-Za-z_À-￿][A-Za-z0-9_.À-￿]*$/.test(name) && !F.parseCell(name) && !/^R\d*C\d*$/i.test(name) && !/^(TRUE|FALSE)$/i.test(name)) return name;
    return "'" + String(name).replace(/'/g, "''") + "'";
  };

  /* functions that Excel files store with the _xlfn. prefix */
  F.FUTURE = new Set(('ACOT ACOTH AGGREGATE ARABIC BASE BETA.DIST BETA.INV BINOM.DIST BINOM.DIST.RANGE BINOM.INV BITAND BITLSHIFT BITOR BITRSHIFT BITXOR CEILING.MATH CEILING.PRECISE CHISQ.DIST CHISQ.DIST.RT CHISQ.INV CHISQ.INV.RT CHISQ.TEST ' +
    'COMBINA CONCAT CONFIDENCE.NORM CONFIDENCE.T COT COTH COVARIANCE.P COVARIANCE.S CSC CSCH DAYS DECIMAL ERF.PRECISE ERFC.PRECISE EXPON.DIST F.DIST F.DIST.RT F.INV F.INV.RT F.TEST FILTERXML FLOOR.MATH FLOOR.PRECISE FORECAST.ETS FORECAST.ETS.CONFINT ' +
    'FORECAST.ETS.SEASONALITY FORECAST.ETS.STAT FORECAST.LINEAR FORMULATEXT GAMMA GAMMA.DIST GAMMA.INV GAMMALN.PRECISE GAUSS HYPGEOM.DIST IFNA IFS IMCOSH IMCOT IMCSC IMCSCH IMSEC IMSECH IMSINH IMTAN ISFORMULA ISOWEEKNUM LOGNORM.DIST LOGNORM.INV ' +
    'MAXIFS MINIFS MODE.MULT MODE.SNGL MUNIT NEGBINOM.DIST NORM.DIST NORM.INV NORM.S.DIST NORM.S.INV NUMBERVALUE PDURATION PERCENTILE.EXC PERCENTILE.INC PERCENTRANK.EXC PERCENTRANK.INC PERMUTATIONA PHI POISSON.DIST QUARTILE.EXC QUARTILE.INC ' +
    'QUERYSTRING RANK.AVG RANK.EQ RRI SEC SECH SHEET SHEETS SKEW.P STDEV.P STDEV.S SWITCH T.DIST T.DIST.2T T.DIST.RT T.INV T.INV.2T T.TEST TEXTJOIN UNICHAR UNICODE VAR.P VAR.S WEBSERVICE WEIBULL.DIST XOR Z.TEST ' +
    'XLOOKUP XMATCH FILTER SORT SORTBY UNIQUE SEQUENCE RANDARRAY LET LAMBDA TEXTBEFORE TEXTAFTER TEXTSPLIT VSTACK HSTACK TOCOL TOROW WRAPROWS WRAPCOLS TAKE DROP CHOOSEROWS CHOOSECOLS EXPAND ARRAYTOTEXT VALUETOTEXT ' +
    'ISOMITTED BYROW BYCOL MAP REDUCE SCAN MAKEARRAY ECMA.CEILING ISO.CEILING NETWORKDAYS.INTL WORKDAY.INTL STOCKHISTORY IMAGE REGEXTEST REGEXEXTRACT REGEXREPLACE GROUPBY PIVOTBY PERCENTOF TRIMRANGE ANCHORARRAY SINGLE FIELDVALUE CONCAT').split(' '));
  F.XLWS = new Set(['FILTER', 'SORT', 'SORTBY']);
  F.storeName = (name) => (F.FUTURE.has(name) ? '_xlfn.' + (F.XLWS.has(name) ? '_xlws.' : '') + name : name);

  const ERRORS = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA', '#SPILL!', '#CALC!', '#FIELD!', '#BLOCKED!', '#UNKNOWN!', '#CONNECT!', '#BUSY!', '#EXTERNAL!', '#PYTHON!'];
  F.ERRORS = ERRORS;

  /* ------------------------------------------------------------ tokens */
  class ParseError extends Error { constructor(msg, pos) { super(msg); this.pos = pos; this.name = 'FormulaError'; } }
  F.ParseError = ParseError;

  const reCell = /^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})(?![A-Za-z0-9_.(\[])/;
  const reCol = /^(\$?)([A-Za-z]{1,3})(?=\s*:)/;
  const reRowPair = /^(\$?)(\d{1,7}):(\$?)(\d{1,7})(?![\d.A-Za-z(])/;
  const reColPair = /^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})(?![A-Za-z0-9_.(\[!])/;
  const reNum = /^(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?/;
  const reIdent = /^[A-Za-z_\\¡-￿][A-Za-z0-9_.\\?¡-￿]*/;
  const reSheetUnq = /^([A-Za-z0-9_.¡-￿]+(?::[A-Za-z0-9_.¡-￿]+)?)!/;

  function tokenize(src) {
    const toks = [];
    let i = 0;
    const n = src.length;
    const prevSignificant = () => { for (let k = toks.length - 1; k >= 0; k--) if (toks[k].t !== 'ws') return toks[k]; return null; };
    const operandEnded = () => { const p = prevSignificant(); return !!p && (p.t === 'num' || p.t === 'str' || p.t === 'bool' || p.t === 'err' || p.t === 'ref' || p.t === 'name' || p.t === 'struct' || p.t === 'spill' || p.t === ')' || p.t === '}' || (p.t === 'op' && p.v === '%')); };
    while (i < n) {
      const ch = src[i];
      const rest = src.slice(i);
      if (ch === ' ' || ch === '\n' || ch === '\r' || ch === '\t' || ch === ' ') { let j = i; while (j < n && /[ \n\r\t ]/.test(src[j])) j++; toks.push({ t: 'ws', v: src.slice(i, j), pos: i }); i = j; continue; }
      if (ch === '"') {
        let j = i + 1, s = '';
        for (;;) {
          if (j >= n) throw new ParseError('Missing closing quotation mark', i);
          if (src[j] === '"') { if (src[j + 1] === '"') { s += '"'; j += 2; continue; } break; }
          s += src[j++];
        }
        toks.push({ t: 'str', v: s, pos: i }); i = j + 1; continue;
      }
      if (ch === '#') {
        /* spill reference: A1# */
        const pv = toks[toks.length - 1];
        if (pv && (pv.t === 'ref' || pv.t === 'name') && !ERRORS.some((x) => rest.toUpperCase().startsWith(x))) { toks.push({ t: 'spill', pos: i }); i++; continue; }
        const e = ERRORS.find((x) => rest.toUpperCase().startsWith(x));
        if (!e) throw new ParseError('Unknown error value', i);
        toks.push({ t: 'err', v: e, pos: i }); i += e.length; continue;
      }
      if (ch === '{' || ch === '}' || ch === '(' || ch === ')' || ch === ',' || ch === ';') { toks.push({ t: ch, pos: i }); i++; continue; }
      if (ch === '<' || ch === '>') { const two = src.slice(i, i + 2); if (two === '<=' || two === '>=' || two === '<>') { toks.push({ t: 'op', v: two, pos: i }); i += 2; } else { toks.push({ t: 'op', v: ch, pos: i }); i++; } continue; }
      if ('+-*/^&=%:@'.includes(ch)) { toks.push({ t: 'op', v: ch, pos: i }); i++; continue; }
      if (ch === '[' && !operandEnded() && !/^\[[^\[\]]+\][^!\[\]]*!/.test(rest)) {
        /* structured reference without a table name: [@Col], [[#This Row],[Col]] */
        const j = matchBracket(src, i);
        toks.push({ t: 'struct', table: null, spec: src.slice(i, j + 1), pos: i }); i = j + 1; continue;
      }
      /* external workbook prefix [n] or [Book.xlsx] */
      let book = null, sheetTxt = null, j0 = i;
      if (ch === '[') { const j = src.indexOf(']', i); if (j > i) { book = src.slice(i + 1, j); j0 = j + 1; } }
      const r2 = src.slice(j0);
      if (r2[0] === "'") {
        let j = j0 + 1, s = '';
        for (;;) {
          if (j >= n) throw new ParseError('Missing closing quote in sheet name', i);
          if (src[j] === "'") { if (src[j + 1] === "'") { s += "'"; j += 2; continue; } break; }
          s += src[j++];
        }
        if (src[j + 1] !== '!') throw new ParseError('Expected ! after sheet name', j);
        sheetTxt = s; j0 = j + 2;
      } else {
        const m = reSheetUnq.exec(r2);
        if (m) { sheetTxt = m[1]; j0 += m[0].length; }
      }
      if (sheetTxt != null || book != null) {
        /* inside a quoted name, "[Book]Sheet" may carry the workbook */
        if (sheetTxt != null && book == null) { const bm = /^(?:.*[\\/])?\[([^\]]+)\](.*)$/.exec(sheetTxt); if (bm) { book = bm[1]; sheetTxt = bm[2]; } }
        let sheet = sheetTxt, sheet2 = null;
        if (sheetTxt != null) { const k = splitSheet3D(sheetTxt); if (k) { sheet = k[0]; sheet2 = k[1]; } }
        const after = src.slice(j0);
        const tok = refToken(after, j0);
        if (tok) { tok.sheet = sheet; if (sheet2) tok.sheet2 = sheet2; if (book != null) tok.book = book; tok.pos = i; toks.push(tok); i = j0 + tok.len; continue; }
        if (after.toUpperCase().startsWith('#REF!')) { toks.push({ t: 'ref', kind: 'err', sheet, book, pos: i }); i = j0 + 5; continue; }
        const im = reIdent.exec(after);
        if (im) { toks.push({ t: 'name', v: im[0], sheet, book, pos: i }); i = j0 + im[0].length; continue; }
        if (book != null && sheetTxt == null) { /* [1]!Name handled above; otherwise fall through */ }
        throw new ParseError('Invalid reference', i);
      }
      const rt = refToken(rest, i);
      if (rt) { rt.pos = i; toks.push(rt); i += rt.len; continue; }
      if (ch === '!') {
        /* "!A1" in a defined name: the sheet the name is used on */
        const bt = refToken(rest.slice(1), i + 1);
        if (bt) { bt.pos = i; bt.bang = true; bt.len += 1; toks.push(bt); i += bt.len; continue; }
      }
      const nm = reNum.exec(rest);
      if (nm) { toks.push({ t: 'num', v: parseFloat(nm[0]), raw: nm[0], pos: i }); i += nm[0].length; continue; }
      const im = reIdent.exec(rest);
      if (im) {
        let id = im[0];
        let k = i + id.length;
        /* names may not end with '.' followed by '(' issues; functions: identifier + '(' */
        let p = k; while (p < n && src[p] === ' ') p++;
        if (src[k] === '(') { toks.push({ t: 'fn', v: id, pos: i }); i = k; continue; }
        if (src[k] === '[') {
          const j = matchBracket(src, k);
          toks.push({ t: 'struct', table: id, spec: src.slice(k, j + 1), pos: i }); i = j + 1; continue;
        }
        const up = id.toUpperCase();
        if (up === 'TRUE' || up === 'FALSE') { toks.push({ t: 'bool', v: up === 'TRUE', pos: i }); i = k; continue; }
        toks.push({ t: 'name', v: id, pos: i }); i = k; void p; continue;
      }
      throw new ParseError('Unexpected character ' + ch, i);
    }
    return toks;
  }
  function matchBracket(src, i) {
    let depth = 0;
    for (let j = i; j < src.length; j++) {
      const c = src[j];
      if (c === "'" && j + 1 < src.length) { j++; continue; } /* escape inside structured refs */
      if (c === '[') depth++;
      else if (c === ']') { depth--; if (depth === 0) return j; }
    }
    throw new ParseError('Missing ] in structured reference', i);
  }
  function splitSheet3D(s) { const k = s.indexOf(':'); if (k <= 0) return null; return [s.slice(0, k), s.slice(k + 1)]; }
  /** cell / area / whole-row / whole-column reference at the start of s */
  function refToken(s, base) {
    let m;
    if ((m = reRowPair.exec(s))) {
      const a = +m[2] - 1, b = +m[4] - 1;
      if (a >= 0 && b >= 0 && a < MAXR && b < MAXR) return { t: 'ref', kind: 'rows', r1: a, r2: b, ra1: !!m[1], ra2: !!m[3], len: m[0].length };
    }
    if ((m = reColPair.exec(s))) {
      const a = F.colIndex(m[2]), b = F.colIndex(m[4]);
      if (a < MAXC && b < MAXC) return { t: 'ref', kind: 'cols', c1: a, c2: b, ca1: !!m[1], ca2: !!m[3], len: m[0].length };
    }
    m = reCell.exec(s);
    if (!m) return null;
    const c = F.colIndex(m[2]), r = +m[4] - 1;
    if (c >= MAXC || r < 0 || r >= MAXR) return null;
    const a = { r, c, ca: !!m[1], ra: !!m[3] };
    const rest = s.slice(m[0].length);
    /* area A1:B2 (whitespace not allowed inside) */
    if (rest[0] === ':') {
      const m2 = reCell.exec(rest.slice(1));
      if (m2) {
        const c2 = F.colIndex(m2[2]), r2 = +m2[4] - 1;
        if (c2 < MAXC && r2 >= 0 && r2 < MAXR) return { t: 'ref', kind: 'area', r1: a.r, c1: a.c, ra1: a.ra, ca1: a.ca, r2, c2, ra2: !!m2[3], ca2: !!m2[1], len: m[0].length + 1 + m2[0].length };
      }
    }
    void base;
    return { t: 'ref', kind: 'cell', r: a.r, c: a.c, ra: a.ra, ca: a.ca, len: m[0].length };
  }

  /* ------------------------------------------------------------ parser */
  const BIN = { ':': 9, ' ': 8, '^': 5, '*': 4, '/': 4, '+': 3, '-': 3, '&': 2, '=': 1, '<': 1, '>': 1, '<=': 1, '>=': 1, '<>': 1 };
  const UNARY_PREC = 7;
  const PCT_PREC = 6;
  /**
   * Parse formula text (without the leading '=') into an AST.
   * opts: {r, c} — the cell the formula belongs to (relative references are stored as absolute positions)
   */
  F.parse = function (text, opts) {
    const toks = tokenize(String(text));
    let p = 0;
    const peek = () => { while (p < toks.length && toks[p].t === 'ws') p++; return toks[p]; };
    const peekRaw = () => toks[p];
    const next = () => { const t = peek(); p++; return t; };
    const expect = (t) => { const k = next(); if (!k || k.t !== t) throw new ParseError(`Expected ${t}`, k ? k.pos : text.length); return k; };
    const startsOperand = (t) => t && (t.t === 'num' || t.t === 'str' || t.t === 'bool' || t.t === 'err' || t.t === 'ref' || t.t === 'name' || t.t === 'fn' || t.t === '(' || t.t === '{' || t.t === 'struct' || (t.t === 'op' && (t.v === '-' || t.v === '+' || t.v === '@')));

    function primary(allowUnion) {
      const t = next();
      if (!t) throw new ParseError('Formula is incomplete', text.length);
      switch (t.t) {
        case 'num': return { t: 'num', v: t.v };
        case 'str': return { t: 'str', v: t.v };
        case 'bool': return { t: 'bool', v: t.v };
        case 'err': return { t: 'err', v: t.v };
        case 'ref': return refNode(t);
        case 'struct': return { t: 'struct', table: t.table, spec: t.spec, sheet: t.sheet };
        case 'name': {
          const nm = t.v.replace(/^_xlpm\./i, '');
          return t.book != null ? { t: 'name', name: nm, sheet: t.sheet, book: t.book } : t.sheet != null ? { t: 'name', name: nm, sheet: t.sheet } : { t: 'name', name: nm };
        }
        case 'fn': {
          const raw = t.v;
          const udf = /^_xludf\./i.test(raw);
          const name = raw.replace(/^_xlfn\./i, '').replace(/^_xlws\./i, '').replace(/^_xludf\./i, '').toUpperCase();
          expect('(');
          let node = { t: 'fn', name, args: callArgs(), raw };
          if (udf) { node.udf = true; node.name = raw.slice(7); }
          /* LAMBDA(x, x*x)(4): immediately invoked */
          while (peekRaw() && peekRaw().t === '(') { next(); node = { t: 'invoke', fn: node, args: callArgs() }; }
          return node;
        }
        case '(': {
          const e = expr(0, true);
          const items = [e];
          let k = peek();
          while (k && k.t === ',') { next(); items.push(expr(0, true)); k = peek(); }
          expect(')');
          return items.length > 1 ? { t: 'paren', a: { t: 'union', items } } : { t: 'paren', a: e };
        }
        case '{': {
          const rows = [[]];
          for (;;) {
            let k = next();
            let neg = false;
            if (k && k.t === 'op' && (k.v === '-' || k.v === '+')) { neg = k.v === '-'; k = next(); }
            if (!k) throw new ParseError('Missing }', text.length);
            let v;
            if ((k.t === ',' || k.t === ';' || k.t === '}') && !neg) {
              /* an empty item (written by some programs): reads as blank */
              rows[rows.length - 1].push({ t: 'missing' });
              if (k.t === '}') break;
              if (k.t === ';') rows.push([]);
              continue;
            }
            if (k.t === 'num') v = { t: 'num', v: neg ? -k.v : k.v };
            else if (k.t === 'str') v = { t: 'str', v: k.v };
            else if (k.t === 'bool') v = { t: 'bool', v: k.v };
            else if (k.t === 'err') v = { t: 'err', v: k.v };
            else throw new ParseError('Array constants can contain only numbers, text, logical values and errors', k.pos);
            rows[rows.length - 1].push(v);
            k = next();
            if (!k) throw new ParseError('Missing }', text.length);
            if (k.t === '}') break;
            if (k.t === ',') continue;
            if (k.t === ';') { rows.push([]); continue; }
            throw new ParseError('Unexpected item in array constant', k.pos);
          }
          const w = rows[0].length;
          if (rows.some((r) => r.length !== w)) throw new ParseError('Array constant rows must have the same length', t.pos);
          return { t: 'arr', rows };
        }
        case 'op':
          if (t.v === '-' || t.v === '+' || t.v === '@') return { t: 'un', op: t.v, a: expr(UNARY_PREC, allowUnion) };
          throw new ParseError('Unexpected operator ' + t.v, t.pos);
        default:
          throw new ParseError('Unexpected ' + (t.t === ')' ? ')' : t.t === ',' ? ',' : t.t), t.pos);
      }
    }
    function callArgs() {
      const args = [];
      let k = peek();
      if (k && k.t === ')') { next(); return args; }
      for (;;) {
        k = peek();
        if (k && (k.t === ',' || k.t === ')')) args.push({ t: 'missing' });
        else args.push(expr(0, false));
        k = next();
        if (!k) throw new ParseError('Missing )', text.length);
        if (k.t === ')') break;
        if (k.t !== ',') throw new ParseError('Expected , or )', k.pos);
        const k2 = peek();
        if (k2 && k2.t === ')') { args.push({ t: 'missing' }); next(); break; }
      }
      return args;
    }
    function refNode(t) {
      const sh = t.sheet != null ? { sheet: t.sheet } : {};
      if (t.bang) sh.bang = true;
      if (t.sheet2) sh.sheet2 = t.sheet2;
      if (t.book != null) sh.book = t.book;
      if (t.kind === 'err') return Object.assign({ t: 'referr' }, sh);
      if (t.kind === 'cell') return Object.assign({ t: 'ref', r: t.r, c: t.c, ra: t.ra, ca: t.ca }, sh);
      if (t.kind === 'area') return Object.assign({ t: 'area', r1: t.r1, c1: t.c1, r2: t.r2, c2: t.c2, ra1: t.ra1, ca1: t.ca1, ra2: t.ra2, ca2: t.ca2 }, sh);
      if (t.kind === 'cols') return Object.assign({ t: 'area', kind: 'cols', r1: 0, r2: MAXR - 1, c1: t.c1, c2: t.c2, ca1: t.ca1, ca2: t.ca2, ra1: true, ra2: true }, sh);
      if (t.kind === 'rows') return Object.assign({ t: 'area', kind: 'rows', c1: 0, c2: MAXC - 1, r1: t.r1, r2: t.r2, ra1: t.ra1, ra2: t.ra2, ca1: true, ca2: true }, sh);
      throw new ParseError('Bad reference', t.pos);
    }
    function expr(minPrec, allowUnion) {
      let left = primary(allowUnion);
      for (;;) {
        /* intersection: whitespace followed by the start of another operand */
        const raw = peekRaw();
        if (raw && raw.t === 'ws') {
          const save = p;
          p++;
          const nx = peekRaw();
          if (nx && startsOperand(nx) && nx.t !== 'op' && BIN[' '] >= minPrec) {
            if (isRefLike(left) && (nx.t === 'ref' || nx.t === 'name' || nx.t === '(' || nx.t === 'fn' || nx.t === 'struct')) {
              const right = expr(BIN[' '] + 1, allowUnion);
              left = { t: 'op', op: ' ', a: left, b: right };
              continue;
            }
          }
          p = save;
        }
        const t = peek();
        if (!t) break;
        if (t.t === 'spill') { next(); left = { t: 'fn', name: 'ANCHORARRAY', args: [left], raw: '_xlfn.ANCHORARRAY' }; continue; }
        if (t.t === 'op' && t.v === '%') { if (PCT_PREC < minPrec) break; next(); left = { t: 'pct', a: left }; continue; }
        if (t.t !== 'op') break;
        const prec = BIN[t.v];
        if (prec == null || prec < minPrec) break;
        next();
        /* all binary operators are left-associative in Excel (2^3^2 = 64) */
        const right = expr(prec + 1, allowUnion);
        left = { t: 'op', op: t.v, a: left, b: right };
      }
      return left;
    }
    const isRefLike = (n) => n && (n.t === 'ref' || n.t === 'area' || n.t === 'name' || n.t === 'paren' || n.t === 'fn' || n.t === 'struct' || (n.t === 'op' && (n.op === ':' || n.op === ' ')));
    if (!toks.filter((t) => t.t !== 'ws').length) throw new ParseError('Empty formula', 0);
    const ast = expr(0, false);
    const rest = peek();
    if (rest) throw new ParseError(rest.t === ')' ? 'Too many closing parentheses' : 'Unexpected text after the formula', rest.pos);
    void opts;
    markParams(ast);
    return ast;
  };
  /* LET / LAMBDA parameter names are written _xlpm.x in files */
  function markParams(ast) {
    F.walk(ast, (n) => {
      if (n.t !== 'fn' || (n.name !== 'LET' && n.name !== 'LAMBDA')) return;
      const params = new Set();
      const last = n.args.length - 1;
      n.args.forEach((a, i) => {
        if (a.t !== 'name') return;
        if ((n.name === 'LET' && i % 2 === 0 && i < last) || (n.name === 'LAMBDA' && i < last)) params.add(a.name.toUpperCase());
      });
      F.walk(n, (k) => { if (k.t === 'name' && k.sheet == null && params.has(k.name.toUpperCase())) k.param = true; });
    });
  }
  F.tokenize = tokenize;
  /** quick check used by the editor */
  F.tryParse = (text, opts) => { try { return { ast: F.parse(text, opts) }; } catch (e) { return { error: e }; } };

  /* --------------------------------------------------------- serialise */
  const fmtNum = (v) => {
    const a = Math.abs(v);
    if (a !== 0 && (a >= 1e15 || a < 1e-5)) return v.toExponential().toUpperCase();
    return String(v);
  };
  const refText = (n, opts) => {
    const pre = sheetPrefix(n, opts);
    if (n.t === 'referr') return pre + '#REF!';
    if (n.t === 'ref') return pre + F.cellName(n.r, n.c, n.ra, n.ca);
    if (n.kind === 'cols') return pre + (n.ca1 ? '$' : '') + F.colName(n.c1) + ':' + (n.ca2 ? '$' : '') + F.colName(n.c2);
    if (n.kind === 'rows') return pre + (n.ra1 ? '$' : '') + (n.r1 + 1) + ':' + (n.ra2 ? '$' : '') + (n.r2 + 1);
    return pre + F.cellName(n.r1, n.c1, n.ra1, n.ca1) + ':' + F.cellName(n.r2, n.c2, n.ra2, n.ca2);
  };
  function sheetPrefix(n, opts) {
    if (n.bang && n.sheet == null) return '!';
    if (n.sheet == null && n.book == null) return '';
    const book = n.book != null ? '[' + n.book + ']' : '';
    if (n.sheet == null) return book + '!';
    const s = n.sheet2 ? n.sheet + ':' + n.sheet2 : n.sheet;
    const q = F.quoteSheet(n.sheet) !== n.sheet || (n.sheet2 && F.quoteSheet(n.sheet2) !== n.sheet2) || (book && /[^A-Za-z0-9_.]/.test(n.book));
    void opts;
    return q ? "'" + (book + s).replace(/'/g, "''") + "'!" : book + s + '!';
  }
  /**
   * AST → formula text (no '=').
   * opts: {store: true} writes the file form (_xlfn. prefixes), otherwise the display form
   */
  F.toText = function (n, opts) {
    opts = opts || {};
    const T = (x) => F.toText(x, opts);
    switch (n.t) {
      case 'num': return fmtNum(n.v);
      case 'str': return '"' + n.v.replace(/"/g, '""') + '"';
      case 'bool': return n.v ? 'TRUE' : 'FALSE';
      case 'err': return n.v;
      case 'missing': return '';
      case 'ref': case 'area': case 'referr': return refText(n, opts);
      case 'name': return (n.sheet != null || n.book != null ? sheetPrefix(n, opts) : '') + (opts.store && n.param ? '_xlpm.' : '') + n.name;
      case 'struct': return (n.table || '') + n.spec;
      case 'fn': {
        if (!opts.store && n.name === 'ANCHORARRAY' && n.args.length === 1 && /^(ref|name|area)$/.test(n.args[0].t)) return T(n.args[0]) + '#';
        if (!opts.store && n.name === 'SINGLE' && n.args.length === 1) return '@' + T(n.args[0]);
        if (n.udf) return (opts.store ? '_xludf.' : '') + n.name + '(' + n.args.map(T).join(',') + ')';
        const nm = opts.store ? (n.raw && /^_xl/i.test(n.raw) && !F.FUTURE.has(n.name) ? n.raw : F.storeName(n.name)) : n.name;
        return nm + '(' + n.args.map(T).join(',') + ')';
      }
      case 'invoke': return T(n.fn) + '(' + n.args.map(T).join(',') + ')';
      case 'paren': return '(' + T(n.a) + ')';
      case 'union': return n.items.map(T).join(',');
      case 'un': return n.op === '@' && opts.store ? '_xlfn.SINGLE(' + T(n.a) + ')' : n.op + T(n.a);
      case 'pct': return T(n.a) + '%';
      case 'op': return T(n.a) + n.op + T(n.b);
      case 'arr': return '{' + n.rows.map((r) => r.map(T).join(',')).join(';') + '}';
      default: return '';
    }
  };

  /* ------------------------------------------------- text forms (spacing kept) */
  /**
   * Rewrite formula text token by token, keeping the author's spacing:
   *  display form — what the formula bar shows (no _xlfn./_xlpm. prefixes, A1# spills, @ for SINGLE)
   *  store form   — what files hold (prefixes added, references and function names upper-cased)
   */
  function rewrite(text, store) {
    let toks;
    try { toks = tokenize(text); } catch (e) { return text; }
    let params = null;
    if (store) {
      params = new Set();
      try { F.walk(F.parse(text), (n) => { if (n.t === 'name' && n.param) params.add(n.name.toUpperCase()); }); } catch (e) { /* keep */ }
    }
    const raw = (k) => text.slice(toks[k].pos, k + 1 < toks.length ? toks[k + 1].pos : text.length);
    const out = [];
    const skip = new Set();
    for (let k = 0; k < toks.length; k++) {
      const t = toks[k];
      const r = raw(k);
      if (skip.has(k)) continue;
      switch (t.t) {
        case 'fn': {
          const bare = t.v.replace(/^_xlfn\./i, '').replace(/^_xlws\./i, '');
          const up = bare.toUpperCase();
          if (/^_xludf\./i.test(t.v)) { out.push(store ? t.v : t.v.slice(7)); break; }
          if (!store && up === 'ANCHORARRAY') {
            /* _xlfn.ANCHORARRAY(A1) → A1# */
            const k1 = k + 1, k2 = k + 2, k3 = k + 3;
            if (toks[k1] && toks[k1].t === '(' && toks[k2] && (toks[k2].t === 'ref' || toks[k2].t === 'name') && toks[k3] && toks[k3].t === ')') { out.push(raw(k2).trim() + '#'); k = k3; break; }
          }
          if (!store && up === 'SINGLE' && toks[k + 1] && toks[k + 1].t === '(') {
            /* _xlfn.SINGLE(x) → @x when x is one operand (reference, name, call, bracketed group), else @(x) */
            let depth = 0, m = -1, simple = true;
            for (let q = k + 1; q < toks.length; q++) {
              const tq = toks[q];
              if (tq.t === '(' || tq.t === 'fn' || tq.t === '{') { if (tq.t !== 'fn') depth++; continue; }
              if (tq.t === ')' || tq.t === '}') { depth--; if (depth === 0) { m = q; break; } continue; }
              if (depth === 1 && (tq.t === 'op' || tq.t === ',' || tq.t === 'sep')) simple = false;
            }
            if (m > 0) {
              if (simple) { out.push('@'); skip.add(m); k = k + 1; } else { out.push('@('); k = k + 1; }
              break;
            }
          }
          out.push(store ? F.storeName(up) : up);
          break;
        }
        case 'spill': {
          if (!store) { out.push('#'); break; }
          /* A1# → _xlfn.ANCHORARRAY(A1) */
          const prev = out.pop();
          out.push('_xlfn.ANCHORARRAY(' + prev + ')');
          break;
        }
        case 'ref': {
          if (t.kind === 'err') { out.push(r.replace(/#ref!/i, '#REF!')); break; }
          out.push(refText(refNodeOf(t), {}));
          break;
        }
        case 'name': {
          const bare = t.v.replace(/^_xlpm\./i, '');
          const pre = t.sheet != null || t.book != null ? sheetPrefix({ sheet: t.sheet, book: t.book }, {}) : '';
          if (store && params && params.has(bare.toUpperCase()) && t.sheet == null) out.push('_xlpm.' + bare);
          else out.push(pre + bare);
          break;
        }
        case 'bool': out.push(t.v ? 'TRUE' : 'FALSE'); break;
        case 'err': out.push(t.v); break;
        default: out.push(r);
      }
    }
    return out.join('');
  }
  function refNodeOf(t) {
    const n = Object.assign({}, t);
    if (t.kind === 'cell') { n.t = 'ref'; }
    else { n.t = 'area'; }
    return n;
  }
  F.display = (text) => rewrite(String(text), false);
  F.toStore = (text) => rewrite(String(text), true);

  /* ------------------------------------------------------------ walking */
  F.walk = function walk(n, fn) {
    if (!n) return;
    fn(n);
    switch (n.t) {
      case 'fn': n.args.forEach((a) => walk(a, fn)); break;
      case 'invoke': walk(n.fn, fn); n.args.forEach((a) => walk(a, fn)); break;
      case 'paren': case 'un': case 'pct': walk(n.a, fn); break;
      case 'op': walk(n.a, fn); walk(n.b, fn); break;
      case 'union': n.items.forEach((a) => walk(a, fn)); break;
      case 'arr': n.rows.forEach((r) => r.forEach((a) => walk(a, fn))); break;
      default: break;
    }
  };
  /** deep copy with a mapping function for reference nodes: map(node) → node | null(keep) */
  F.map = function map(n, fn) {
    const r = fn(n);
    if (r) return r;
    switch (n.t) {
      case 'fn': return Object.assign({}, n, { args: n.args.map((a) => map(a, fn)) });
      case 'invoke': return Object.assign({}, n, { fn: map(n.fn, fn), args: n.args.map((a) => map(a, fn)) });
      case 'paren': case 'un': case 'pct': return Object.assign({}, n, { a: map(n.a, fn) });
      case 'op': return Object.assign({}, n, { a: map(n.a, fn), b: map(n.b, fn) });
      case 'union': return Object.assign({}, n, { items: n.items.map((a) => map(a, fn)) });
      default: return Object.assign({}, n);
    }
  };

  /* --------------------------------------------------- copy / fill shift */
  /** shift relative parts by (dr, dc); refs pushed off the sheet become #REF! */
  /** move relative references by (dr, dc); wrap: offsets wrap around the sheet edges (relative defined names) */
  F.shift = function (ast, dr, dc, wrap) {
    if (!dr && !dc) return ast;
    const wr = (v) => (wrap ? ((v % MAXR) + MAXR) % MAXR : v), wc = (v) => (wrap ? ((v % MAXC) + MAXC) % MAXC : v);
    return F.map(ast, (n) => {
      if (n.t === 'ref') {
        const r = n.ra ? n.r : wr(n.r + dr), c = n.ca ? n.c : wc(n.c + dc);
        if (r < 0 || r >= MAXR || c < 0 || c >= MAXC) return refErr(n);
        return Object.assign({}, n, { r, c });
      }
      if (n.t === 'area') {
        const o = Object.assign({}, n);
        if (n.kind !== 'cols') { o.r1 = n.ra1 ? n.r1 : wr(n.r1 + dr); o.r2 = n.ra2 ? n.r2 : wr(n.r2 + dr); }
        if (n.kind !== 'rows') { o.c1 = n.ca1 ? n.c1 : wc(n.c1 + dc); o.c2 = n.ca2 ? n.c2 : wc(n.c2 + dc); }
        if (o.r1 < 0 || o.r2 >= MAXR || o.c1 < 0 || o.c2 >= MAXC || o.r2 < 0 || o.c2 < 0 || o.r1 >= MAXR || o.c1 >= MAXC) return refErr(n);
        return o;
      }
      return null;
    });
  };
  const refErr = (n) => { const o = { t: 'referr' }; if (n.sheet != null) o.sheet = n.sheet; if (n.sheet2) o.sheet2 = n.sheet2; if (n.book != null) o.book = n.book; return o; };
  /** translate formula text written for (r0,c0) to (r1,c1) */
  F.translate = function (text, dr, dc) {
    try { return F.toText(F.shift(F.parse(text), dr, dc)); } catch (e) { return text; }
  };

  /* ------------------------------------------- insert / delete rows & columns */
  const sameSheet = (n, target, home) => {
    const s = n.sheet != null ? n.sheet : home;
    if (n.book != null) return false;
    if (n.sheet2) {
      return false; /* 3-D references keep their positions (handled conservatively) */
    }
    return s != null && target != null && String(s).toLowerCase() === String(target).toLowerCase();
  };
  /**
   * Adjust references for a structural change on sheet `op.sheet`:
   * op = {sheet, axis:'r'|'c', at, n} — n > 0 inserts n rows/cols before `at`; n < 0 deletes -n starting at `at`.
   * home = the sheet the formula lives on (for unqualified refs).
   * Returns {ast, changed}
   */
  F.adjust = function (ast, op, home) {
    let changed = false;
    const ins = op.n > 0;
    const cnt = Math.abs(op.n);
    const lim = op.axis === 'r' ? MAXR : MAXC;
    const mv = (x) => {
      if (ins) return x >= op.at ? x + cnt : x;
      if (x < op.at) return x;
      if (x >= op.at + cnt) return x - cnt;
      return null; /* deleted */
    };
    const out = F.map(ast, (n) => {
      if ((n.t !== 'ref' && n.t !== 'area') || !sameSheet(n, op.sheet, home)) return null;
      if (n.t === 'ref') {
        const k = op.axis === 'r' ? 'r' : 'c';
        const v = mv(n[k]);
        if (v == null || v >= lim) { changed = true; return refErr(n); }
        if (v !== n[k]) { changed = true; return Object.assign({}, n, { [k]: v }); }
        return null;
      }
      const k1 = op.axis === 'r' ? 'r1' : 'c1', k2 = op.axis === 'r' ? 'r2' : 'c2';
      if ((op.axis === 'r' && n.kind === 'cols') || (op.axis === 'c' && n.kind === 'rows')) return null;
      let a = n[k1], b = n[k2];
      if (ins) {
        if (a >= op.at) a += cnt;
        if (b >= op.at) b += cnt; /* insertion inside or just after the range's first row widens it */
        if (b >= lim) b = lim - 1;
        if (a >= lim) { changed = true; return refErr(n); }
      } else {
        const d0 = op.at, d1 = op.at + cnt - 1;
        if (b < d0) return null;
        if (a > d1) { a -= cnt; b -= cnt; }
        else if (a >= d0 && b <= d1) { changed = true; return refErr(n); }
        else {
          /* partial overlap: shrink */
          const na = a < d0 ? a : d0;
          const nb = b > d1 ? b - cnt : d0 - 1;
          a = na; b = nb;
        }
      }
      if (a !== n[k1] || b !== n[k2]) { changed = true; return Object.assign({}, n, { [k1]: a, [k2]: b }); }
      return null;
    });
    return { ast: changed ? out : ast, changed };
  };
  /** references into `src` (a range on sheet `sheet`) follow a cut-and-paste move by (dr, dc) to sheet `toSheet` */
  F.adjustMove = function (ast, src, dr, dc, home, toSheet) {
    let changed = false;
    const inside = (r, c) => r >= src.r1 && r <= src.r2 && c >= src.c1 && c <= src.c2;
    const out = F.map(ast, (n) => {
      if ((n.t !== 'ref' && n.t !== 'area') || !sameSheet(n, src.sheet, home)) return null;
      const qual = (o) => { if (toSheet && String(toSheet).toLowerCase() !== String(src.sheet).toLowerCase()) o.sheet = toSheet; return o; };
      if (n.t === 'ref' && inside(n.r, n.c)) { changed = true; return qual(Object.assign({}, n, { r: n.r + dr, c: n.c + dc })); }
      if (n.t === 'area' && !n.kind && inside(n.r1, n.c1) && inside(n.r2, n.c2)) { changed = true; return qual(Object.assign({}, n, { r1: n.r1 + dr, r2: n.r2 + dr, c1: n.c1 + dc, c2: n.c2 + dc })); }
      return null;
    });
    return { ast: changed ? out : ast, changed };
  };
  /** sheet renamed (or deleted when to == null → #REF!) */
  F.renameSheet = function (ast, from, to) {
    let changed = false;
    const f = String(from).toLowerCase();
    const out = F.map(ast, (n) => {
      if (n.t !== 'ref' && n.t !== 'area' && n.t !== 'referr' && n.t !== 'name') return null;
      if (n.sheet == null || n.book != null) return null;
      const hit1 = String(n.sheet).toLowerCase() === f, hit2 = n.sheet2 && String(n.sheet2).toLowerCase() === f;
      if (!hit1 && !hit2) return null;
      changed = true;
      if (to == null) return n.t === 'name' ? Object.assign({}, n) : refErr({ sheet: n.sheet });
      const o = Object.assign({}, n);
      if (hit1) o.sheet = to;
      if (hit2) o.sheet2 = to;
      return o;
    });
    return { ast: changed ? out : ast, changed };
  };
  /** F4: cycle $ on the reference at caret position in editing text. Returns {text, caret} */
  F.toggleAbs = function (text, selStart, selEnd) {
    const re = /((?:'[^']*'|[A-Za-z0-9_.]+)!)?(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})(?::(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7}))?/g;
    let m;
    while ((m = re.exec(text))) {
      const s = m.index, e = s + m[0].length;
      if (selEnd < s || selStart > e) continue;
      const cyc = (ca, ra) => (ca && ra ? ['', ''] : !ca && ra ? ['$', ''] : ca && !ra ? ['', ''] : ['$', '$']);
      /* Excel cycle: A1 → $A$1 → A$1 → $A1 → A1 */
      const next = (ca, ra) => (!ca && !ra ? ['$', '$'] : ca && ra ? ['', '$'] : !ca && ra ? ['$', ''] : ['', '']);
      void cyc;
      const [c1, r1] = next(!!m[2], !!m[4]);
      let rep = (m[1] || '') + c1 + m[3] + r1 + m[5];
      if (m[7]) { const [c2, r2] = next(!!m[6], !!m[8]); rep += ':' + c2 + m[7] + r2 + m[9]; }
      return { text: text.slice(0, s) + rep + text.slice(e), caret: s + rep.length, start: s };
    }
    return null;
  };
})(typeof window !== 'undefined' ? window : globalThis);
