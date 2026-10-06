/* VibeOffice — number formats (Ledger's cells, Lectern's chart labels).
 * Excel format codes (sections, conditions, colours, dates, elapsed time, fractions, scientific, text),
 * the built-in format table, the General format, and parsing of typed input into values.
 * Works in the browser and in Node (tests): attaches to the global L.
 */
(function (root) {
  'use strict';
  const L = root.L || (root.L = {});
  const NF = (L.numfmt = {});

  NF.BUILTIN = {
    0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00',
    5: '"$"#,##0_);\\("$"#,##0\\)', 6: '"$"#,##0_);[Red]\\("$"#,##0\\)', 7: '"$"#,##0.00_);\\("$"#,##0.00\\)', 8: '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',
    9: '0%', 10: '0.00%', 11: '0.00E+00', 12: '# ?/?', 13: '# ??/??',
    14: 'm/d/yyyy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss', 22: 'm/d/yyyy h:mm',
    37: '#,##0 ;(#,##0)', 38: '#,##0 ;[Red](#,##0)', 39: '#,##0.00;(#,##0.00)', 40: '#,##0.00;[Red](#,##0.00)',
    41: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)', 42: '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)',
    43: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)', 44: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)',
    45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mmss.0', 48: '##0.0E+0', 49: '@',
    /* East Asian ids that en-US Excel shows with these codes */
    27: 'm/d/yyyy', 28: 'm/d/yyyy', 29: 'm/d/yyyy', 30: 'm/d/yy', 31: 'm/d/yyyy', 32: 'h:mm:ss', 33: 'h:mm:ss', 34: 'h:mm:ss', 35: 'h:mm:ss', 36: 'm/d/yyyy',
    50: 'm/d/yyyy', 51: 'm/d/yyyy', 52: 'm/d/yyyy', 53: 'm/d/yyyy', 54: 'm/d/yyyy', 55: 'm/d/yyyy', 56: 'm/d/yyyy', 57: 'm/d/yyyy', 58: 'm/d/yyyy',
    59: 't0', 60: 't0.00', 61: 't#,##0', 62: 't#,##0.00', 67: 't0%', 68: 't0.00%', 69: 't# ?/?', 70: 't# ??/??',
  };
  NF.builtinId = (code) => { for (const k in NF.BUILTIN) if (NF.BUILTIN[k] === code && +k < 50) return +k; return -1; };

  const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  NF.MONTHS = MONTHS; NF.DAYS = DAYS;
  const COLORS = { black: '#000000', blue: '#0000FF', cyan: '#00FFFF', green: '#00FF00', magenta: '#FF00FF', red: '#FF0000', white: '#FFFFFF', yellow: '#FFFF00' };
  /* the 56-colour legacy palette ([ColorN] and indexed colours) */
  NF.PALETTE = ['#000000', '#FFFFFF', '#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#FF00FF', '#00FFFF', '#800000', '#008000', '#000080', '#808000', '#800080', '#008080', '#C0C0C0', '#808080',
    '#9999FF', '#993366', '#FFFFCC', '#CCFFFF', '#660066', '#FF8080', '#0066CC', '#CCCCFF', '#000080', '#FF00FF', '#FFFF00', '#00FFFF', '#800080', '#800000', '#008080', '#0000FF',
    '#00CCFF', '#CCFFFF', '#CCFFCC', '#FFFF99', '#99CCFF', '#FF99CC', '#CC99FF', '#FFCC99', '#3366FF', '#33CCCC', '#99CC00', '#FFCC00', '#FF9900', '#FF6600', '#666699', '#969696',
    '#003366', '#339966', '#003300', '#333300', '#993300', '#993366', '#333399', '#333333'];

  /* ------------------------------------------------------------------ dates */
  /** serial → {y, m (1-12), d, wd (0=Sunday), H, M, S, frac(sub-second 0..1)} ; date1904 flag */
  NF.serialToParts = function (v, date1904, subDigits) {
    let days = Math.floor(v);
    let secs = (v - days) * 86400;
    const scale = Math.pow(10, subDigits || 0);
    secs = Math.round(secs * scale) / scale;
    /* Excel rounds tiny binary noise: 0.9999999999 days is midnight of the next day */
    if (secs >= 86400) { secs -= 86400; days += 1; }
    let y, m, d, wd;
    if (date1904) {
      const dt = new Date(Date.UTC(1904, 0, 1) + days * 86400000);
      y = dt.getUTCFullYear(); m = dt.getUTCMonth() + 1; d = dt.getUTCDate(); wd = dt.getUTCDay();
    } else if (days === 0) { y = 1900; m = 1; d = 0; wd = 6; }
    else if (days === 60) { y = 1900; m = 2; d = 29; wd = 3; }
    else {
      const dt = new Date(Date.UTC(1899, 11, days < 60 ? 31 : 30) + days * 86400000);
      y = dt.getUTCFullYear(); m = dt.getUTCMonth() + 1; d = dt.getUTCDate();
      wd = ((days % 7) + 6) % 7; /* Excel's weekday: serial 1 is a Sunday */
    }
    const whole = Math.floor(secs);
    const H = Math.floor(whole / 3600), M = Math.floor((whole % 3600) / 60), S = whole % 60;
    return { y, m, d, wd, H, M, S, sub: secs - whole, days, secs };
  };
  /** y, m (1-based, may overflow), d → serial (1900 system with the 1900-02-29 quirk) */
  NF.dateToSerial = function (y, m, d, date1904) {
    const t = Date.UTC(y, m - 1, 1) + (d - 1) * 86400000;
    if (date1904) return Math.round((t - Date.UTC(1904, 0, 1)) / 86400000);
    let s = Math.round((t - Date.UTC(1899, 11, 30)) / 86400000);
    if (s <= 60) s -= 1; /* before 1900-03-01 the real calendar is one day behind Excel's */
    if (y === 1900 && m === 2 && d === 29) return 60;
    return s;
  };
  NF.jsDateToSerial = (dt, date1904) => NF.dateToSerial(dt.getFullYear(), dt.getMonth() + 1, dt.getDate(), date1904) + (dt.getHours() * 3600 + dt.getMinutes() * 60 + dt.getSeconds() + dt.getMilliseconds() / 1000) / 86400;
  NF.serialToJSDate = (v, date1904) => { const p = NF.serialToParts(v, date1904, 3); return new Date(p.y, p.m - 1, p.d, p.H, p.M, p.S, Math.round(p.sub * 1000)); };

  /* ------------------------------------------------------------- rounding */
  /** decimal string of |x| rounded half away from zero to `dec` places, Excel style (15 significant digits first) */
  function fixed(x, dec) {
    x = Math.abs(x);
    if (!isFinite(x)) return 'NaN';
    if (x === 0) return dec > 0 ? '0.' + '0'.repeat(dec) : '0';
    /* 15 significant digits, as Excel stores them */
    let s = x.toPrecision(15);
    let mant, exp;
    const ei = s.indexOf('e');
    if (ei >= 0) { mant = s.slice(0, ei); exp = parseInt(s.slice(ei + 1), 10); } else { mant = s; exp = 0; }
    let [ip, fp] = mant.split('.');
    fp = fp || '';
    let digits = ip + fp; /* decimal point after ip.length (+exp) */
    let point = ip.length + exp;
    digits = digits.replace(/^0+/, (z) => { point -= z.length; return ''; });
    if (!digits) return dec > 0 ? '0.' + '0'.repeat(dec) : '0';
    /* now value = 0.digits × 10^point */
    const keep = point + dec; /* digits to keep */
    if (keep < 0) return dec > 0 ? '0.' + '0'.repeat(dec) : '0';
    let arr = digits.split('').map(Number);
    if (arr.length > keep) {
      const up = arr[keep] >= 5;
      arr = arr.slice(0, keep);
      if (up) {
        let i = arr.length - 1;
        while (i >= 0) { if (arr[i] === 9) { arr[i] = 0; i--; } else { arr[i]++; break; } }
        if (i < 0) { arr.unshift(1); point++; }
      }
    }
    while (arr.length < point + dec) arr.push(0);
    let intPart, fracPart;
    if (point <= 0) { intPart = '0'; fracPart = '0'.repeat(-point) + arr.join(''); }
    else { intPart = arr.slice(0, point).join('') || '0'; fracPart = arr.slice(point).join(''); }
    fracPart = (fracPart + '0'.repeat(dec)).slice(0, dec);
    intPart = intPart.replace(/^0+(?=\d)/, '');
    return dec > 0 ? intPart + '.' + fracPart : intPart;
  }
  NF.fixed = fixed;
  /** round to 15 significant digits (what Excel displays and compares) */
  NF.r15 = (x) => (x === 0 || !isFinite(x) ? x : parseFloat(x.toPrecision(15)));

  /* -------------------------------------------------------------- General */
  const stripDec = (o) => (o.indexOf('.') < 0 || /e/i.test(o) ? o : o.replace(/(?:\.0*|(\.\d*[1-9])0+)$/, '$1'));
  const normExp = (o) => (o.indexOf('E') < 0 ? o : o.replace(/(?:\.0*|(\.\d*[1-9])0+)E/, '$1E').replace(/(E[+-])(\d)$/, '$10$2'));
  /** General format in at most `w` characters (Excel's default is 11; the grid passes the column width) */
  function sciForm(a, V, width) {
    const expPart = Math.abs(V) >= 100 ? 5 : 4; /* E+XX or E+XXX */
    for (let d = Math.max(0, width - 2 - expPart); d >= 0; d--) {
      let s = a.toExponential(d).toUpperCase();
      s = normExp(s.replace(/E([+-])(\d+)$/, (m0, sg, e) => 'E' + sg + (e.length < 2 ? '0' + e : e)));
      if (s.length <= width || d === 0) return s;
    }
    return null;
  }
  /** General format. w: width in characters (11 for a standard cell). fn: the TEXT() variant, which keeps small numbers in decimal form when that loses no digits */
  NF.general = function (v, w, fn) {
    if (typeof v !== 'number') return String(v);
    if (v === 0) return '0';
    if (!isFinite(v)) return '#NUM!';
    w = w || 11;
    const neg = v < 0;
    const width = w; /* the minus sign does not count (Excel gives negatives one more character) */
    const a = Math.abs(v);
    const V = Math.floor(Math.log10(a));
    let out = null;
    if (V + 1 <= width && V >= (fn ? -(width - 2) : -4)) {
      /* plain decimal: as many digits as fit */
      const intLen = V >= 0 ? V + 1 : 1;
      let dec = Math.max(0, width - intLen - 1);
      for (; dec >= 0; dec--) {
        let s = stripDec(fixed(a, Math.min(dec, 30)));
        if (s.length <= width) {
          if (s === '0' || /^0\.?0*$/.test(s)) { s = null; }
          out = s;
          break;
        }
      }
      if (out != null && V < 0) {
        /* small numbers: decimal only when it keeps as many significant digits as the scientific form */
        const sig = out.replace(/^0\.0*/, '').length;
        const sci = sciForm(a, V, width);
        const sigSci = sci ? sci.replace(/E.*$/, '').replace('.', '').length : 0;
        if (sig < 1 || (fn && sig < sigSci)) out = null;
      }
    }
    if (out == null) {
      /* scientific: mantissa digits that fit */
      const expPart = Math.abs(V) >= 100 ? 5 : 4; /* E+XX or E+XXX */
      for (let d = Math.max(0, width - 2 - expPart); d >= 0; d--) {
        let s = a.toExponential(d).toUpperCase();
        s = normExp(s.replace(/E([+-])(\d+)$/, (m0, sg, e) => 'E' + sg + (e.length < 2 ? '0' + e : e)));
        if (s.length <= width || d === 0) { out = s; break; }
      }
      if (out.length > width) return '#'.repeat(Math.max(1, w));
    }
    return (neg ? '-' : '') + out;
  };

  /* --------------------------------------------------------------- parser */
  const cache = new Map();
  /** split a code into sections at ';' outside quotes, brackets and escapes */
  function splitSections(code) {
    const out = [];
    let cur = '', q = false, b = false;
    for (let i = 0; i < code.length; i++) {
      const ch = code[i];
      if (q) { cur += ch; if (ch === '"') q = false; continue; }
      if (ch === '\\' || ch === '_' || ch === '*') { cur += ch + (code[i + 1] || ''); i++; continue; }
      if (ch === '"') { q = true; cur += ch; continue; }
      if (ch === '[') b = true;
      if (ch === ']') b = false;
      if (ch === ';' && !b) { out.push(cur); cur = ''; continue; }
      cur += ch;
    }
    out.push(cur);
    return out;
  }
  const DATE_RE = /^(yyyy|yyy|yy|y|mmmmm|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|e|bbbb|bb|b|gggg|ggg|gg|g)/i;
  /** tokenise one section */
  function tokenize(sec) {
    const toks = [];
    const lit = (s) => { const last = toks[toks.length - 1]; if (last && last.t === 'lit') last.s += s; else toks.push({ t: 'lit', s }); };
    const info = { color: null, cond: null, currency: null, lcid: null, elapsed: false, date: false, text: false, general: false, num: false, fill: false };
    let i = 0;
    const n = sec.length;
    while (i < n) {
      const ch = sec[i];
      const rest = sec.slice(i);
      if (ch === '"') { const j = sec.indexOf('"', i + 1); const s = j < 0 ? sec.slice(i + 1) : sec.slice(i + 1, j); lit(s); i = j < 0 ? n : j + 1; continue; }
      if (ch === '\\') { lit(sec[i + 1] || ''); i += 2; continue; }
      if (ch === '_') { toks.push({ t: 'pad', ch: sec[i + 1] || ' ' }); i += 2; continue; }
      if (ch === '*') { toks.push({ t: 'fill', ch: sec[i + 1] || ' ' }); info.fill = true; i += 2; continue; }
      if (ch === '[') {
        const j = sec.indexOf(']', i);
        const body = j < 0 ? sec.slice(i + 1) : sec.slice(i + 1, j);
        i = j < 0 ? n : j + 1;
        const lb = body.toLowerCase();
        if (COLORS[lb]) { info.color = COLORS[lb]; continue; }
        let m;
        if ((m = /^color\s*(\d+)$/.exec(lb))) { info.color = NF.PALETTE[(+m[1] - 1) % 56] || null; continue; }
        if ((m = /^(<=|>=|<>|<|>|=)\s*(-?[\d.]+(?:e[+-]?\d+)?)$/.exec(lb))) { info.cond = { op: m[1], v: parseFloat(m[2]) }; continue; }
        if (/^(h+|m+|s+)$/.test(lb)) { toks.push({ t: 'el', u: lb[0], len: lb.length }); info.elapsed = true; info.date = true; continue; }
        if (body[0] === '$') {
          const dash = body.indexOf('-', 1);
          const sym = dash < 0 ? body.slice(1) : body.slice(1, dash);
          const lcid = dash < 0 ? '' : body.slice(dash + 1);
          info.lcid = lcid.toUpperCase();
          if (sym) lit(sym);
          continue;
        }
        if (/^dbnum\d$|^natnum\d$/.test(lb)) continue;
        continue; /* unknown bracket: ignore */
      }
      if (/^general/i.test(rest)) { toks.push({ t: 'gen' }); info.general = true; i += 7; continue; }
      if (/^(am\/pm|a\/p)/i.test(rest)) { const m = /^(am\/pm|a\/p)/i.exec(rest)[0]; toks.push({ t: 'ampm', s: m }); info.date = true; i += m.length; continue; }
      const dm = DATE_RE.exec(rest);
      if (dm && !(ch === 'e' || ch === 'E') && !(ch === 'g' || ch === 'G') && !(ch === 'b' || ch === 'B')) {
        toks.push({ t: 'dt', s: dm[0].toLowerCase() }); info.date = true; i += dm[0].length; continue;
      }
      if (dm && (ch === 'e' || ch === 'E')) {
        /* E+ / E- is scientific notation in a number section; a bare e is the (Gregorian) year */
        if (sec[i + 1] === '+' || sec[i + 1] === '-') { toks.push({ t: 'exp', sign: sec[i + 1], upper: ch === 'E' }); info.num = true; i += 2; continue; }
        toks.push({ t: 'dt', s: 'yyyy' }); info.date = true; i++; continue;
      }
      if (dm && (ch === 'b' || ch === 'B')) { toks.push({ t: 'dt', s: dm[0].length > 2 ? 'bbbb' : 'bb' }); info.date = true; i += dm[0].length; continue; }
      if (dm && (ch === 'g' || ch === 'G')) { i += dm[0].length; info.date = true; continue; }
      if (ch === '0' || ch === '#' || ch === '?') { toks.push({ t: 'd', ch }); info.num = true; i++; continue; }
      if (ch === '.') {
        /* fractional seconds after s */
        if (info.date) { const m = /^\.(0+)/.exec(rest); if (m) { toks.push({ t: 'sub', len: m[1].length }); i += m[0].length; continue; } }
        toks.push({ t: '.' }); i++; continue;
      }
      if (ch === ',') { toks.push({ t: ',' }); i++; continue; }
      if (ch === '%') { toks.push({ t: '%' }); info.num = true; i++; continue; }
      if (ch === '/') { toks.push({ t: '/' }); i++; continue; }
      if (ch === '@') { toks.push({ t: '@' }); info.text = true; i++; continue; }
      lit(ch); i++;
    }
    return { toks, info };
  }

  /** parse a full code → {sections:[...]} */
  NF.parse = function (code) {
    code = code == null ? 'General' : String(code);
    let p = cache.get(code);
    if (p) return p;
    const raw = splitSections(code);
    const secs = raw.map((s) => {
      const { toks, info } = tokenize(s);
      const sec = Object.assign({ toks, src: s }, info);
      sec.kind = sec.date && !sec.num ? 'date' : sec.general ? 'general' : sec.num ? 'num' : sec.text ? 'text' : sec.date ? 'date' : 'lit';
      if (sec.kind === 'date' && /^F[48]00$/.test(sec.lcid || '')) {
        /* system long date / time formats */
        const sys = sec.lcid === 'F800' ? 'dddd, mmmm d, yyyy' : 'h:mm:ss AM/PM';
        const t2 = tokenize(sys);
        sec.toks = t2.toks;
      }
      if (sec.kind === 'date') markMinutes(sec.toks);
      if (sec.kind === 'num') analyseNumber(sec);
      return sec;
    });
    p = { code, secs, isDate: secs[0] && secs[0].kind === 'date', isText: secs.length === 1 && secs[0].kind === 'text', hasCond: secs.some((s) => s.cond) };
    if (cache.size > 5000) cache.clear();
    cache.set(code, p);
    return p;
  };
  /* 'm' means minutes right after an hour or right before a second */
  function markMinutes(toks) {
    const dts = toks.filter((t) => t.t === 'dt' || t.t === 'el' || t.t === 'ampm');
    for (let k = 0; k < dts.length; k++) {
      const t = dts[k];
      if (t.t !== 'dt' || (t.s !== 'm' && t.s !== 'mm')) continue;
      const prev = dts[k - 1], next = dts[k + 1];
      if ((prev && ((prev.t === 'dt' && (prev.s[0] === 'h' || prev.s[0] === 's')) || (prev.t === 'el' && (prev.u === 'h' || prev.u === 's')))) || (next && ((next.t === 'dt' && next.s[0] === 's') || (next.t === 'el' && next.u === 's')))) t.min = true;
    }
    /* once a unit is shown as elapsed ([h], [m], [s]), every plain token of that unit is elapsed too */
    const el = new Set(toks.filter((t) => t.t === 'el').map((t) => t.u));
    if (el.size) {
      for (let k = 0; k < toks.length; k++) {
        const t = toks[k];
        if (t.t !== 'dt') continue;
        const u = t.s[0] === 'h' ? 'h' : t.s[0] === 's' ? 's' : t.min ? 'm' : null;
        if (u && el.has(u)) toks[k] = { t: 'el', u, len: t.s.length };
      }
    }
  }
  /* number section layout: integer / decimal / exponent / fraction placeholders, scaling and grouping */
  function analyseNumber(sec) {
    const toks = sec.toks;
    let pct = 0;
    for (const t of toks) if (t.t === '%') pct++;
    sec.pct = pct;
    const ei = toks.findIndex((t) => t.t === 'exp');
    const si = toks.findIndex((t) => t.t === '/');
    sec.frac = si >= 0 && toks.slice(0, si).some((t) => t.t === 'd') && toks.slice(si).some((t) => t.t === 'd' || (t.t === 'lit' && /^\d+$/.test(t.s.trim())));
    if (sec.frac) {
      /* fraction: [int] [num]/[den] */
      const before = toks.slice(0, si), after = toks.slice(si + 1);
      /* numerator placeholders: last run of digits before '/', integer part: an earlier run separated by a literal/space */
      let k = before.length - 1;
      while (k >= 0 && before[k].t !== 'd') k--;
      let numEnd = k;
      while (k >= 0 && (before[k].t === 'd' || before[k].t === ',')) k--;
      const numStart = k + 1;
      sec.fracNum = before.slice(numStart, numEnd + 1).filter((t) => t.t === 'd');
      sec.fracInt = before.slice(0, numStart).some((t) => t.t === 'd');
      sec.fracIntToks = before.slice(0, numStart);
      sec.fracMid = before.slice(numEnd + 1); /* literals between numerator and slash */
      sec.fracNumToks = before.slice(numStart, numEnd + 1);
      let fixedDen = null;
      const denToks = [];
      let j = 0;
      for (; j < after.length; j++) {
        const t = after[j];
        if (t.t === 'd') denToks.push(t);
        else if (t.t === 'lit' && /^\d+/.test(t.s) && !denToks.length) { const m = /^(\d+)(.*)$/.exec(t.s); fixedDen = +m[1]; if (m[2]) { after.splice(j + 1, 0, { t: 'lit', s: m[2] }); } j++; break; }
        else if (denToks.length) break;
        else if (t.t === 'lit' || t.t === 'pad') continue; else break;
      }
      sec.fracDen = denToks;
      sec.fixedDen = fixedDen && fixedDen > 0 ? fixedDen : null;
      sec.fracTail = after.slice(fixedDen ? j : after.indexOf(denToks[denToks.length - 1]) + 1);
      sec.fracPre = after.slice(0, fixedDen ? Math.max(0, j - 1) : after.indexOf(denToks[0]));
      return;
    }
    const end = ei >= 0 ? ei : toks.length;
    const di = toks.findIndex((t, i) => t.t === '.' && i < end);
    const intEnd = di >= 0 ? di : end;
    /* thousands grouping: a comma between integer digit placeholders; trailing commas scale by 1000 */
    let firstD = -1, lastD = -1;
    for (let i = 0; i < intEnd; i++) if (toks[i].t === 'd') { if (firstD < 0) firstD = i; lastD = i; }
    let group = false, scale = 0;
    for (let i = 0; i < intEnd; i++) {
      if (toks[i].t !== ',') continue;
      if (i > firstD && i < lastD) group = true;
      else if (i > lastD && firstD >= 0) {
        /* commas right after the last integer digit (and before '.' or end) scale */
        let allCommas = true;
        for (let k = lastD + 1; k < i; k++) if (toks[k].t !== ',') allCommas = false;
        if (allCommas) scale++; else toks[i] = { t: 'lit', s: ',', bare: true };
      } else toks[i] = { t: 'lit', s: ',', bare: true };
    }
    if (di >= 0) {
      for (let i = di + 1; i < end; i++) if (toks[i].t === ',') { /* commas right after decimals also scale */
        let only = true; for (let k = i + 1; k < end; k++) if (toks[k].t === 'd') only = false;
        if (only) { scale++; toks[i] = { t: 'skip' }; } else toks[i] = { t: 'lit', s: ',', bare: true };
      }
    }
    if (firstD < 0 && di >= 0) { /* ".00" style: no integer placeholders */ }
    /* in scientific formats commas neither group nor print */
    if (ei >= 0) for (let i = 0; i < toks.length; i++) if (toks[i].bare || (i > ei && toks[i].t === ',')) toks[i] = { t: 'skip' };
    sec.group = group;
    sec.scale = scale;
    sec.intEnd = intEnd;
    sec.decAt = di;
    sec.expAt = ei;
    sec.intD = toks.slice(0, intEnd).filter((t) => t.t === 'd').length;
    sec.decD = di >= 0 ? toks.slice(di + 1, end).filter((t) => t.t === 'd').length : 0;
    if (ei >= 0) sec.expD = toks.slice(ei + 1).filter((t) => t.t === 'd').length;
  }

  /* ------------------------------------------------------------- choosing */
  function condOK(c, v) {
    switch (c.op) {
      case '<': return v < c.v; case '>': return v > c.v; case '=': return v === c.v;
      case '<=': return v <= c.v; case '>=': return v >= c.v; case '<>': return v !== c.v;
      default: return false;
    }
  }
  const GENERAL_SEC = { kind: 'general', toks: [{ t: 'gen' }], color: null };
  /** pick the section for a number; returns {sec, abs} where abs = drop the sign */
  function pickNum(p, v) {
    const s = p.secs;
    const nonText = s.filter((x) => !(x.kind === 'text' && s.length > 1 && x === s[3]));
    if (p.hasCond) {
      const c0 = s[0] && s[0].cond, c1 = s[1] && s[1].cond;
      if (c0 && condOK(c0, v)) return { sec: s[0], abs: false };
      if (s[1] && c1 && condOK(c1, v)) return { sec: s[1], abs: v < 0 && c1.op[0] === '<' && c1.v <= 0 };
      if (s[1] && !c1 && !c0) return { sec: s[1], abs: v < 0 };
      if (c0 && !c1 && s[1] && s[1].kind !== 'text') {
        /* "[cond]a;b" — b covers everything else */
        if (s.length >= 3 && v < 0 && s[2] && s[2].kind !== 'text') return { sec: s[1], abs: true };
        return { sec: s[1], abs: false };
      }
      if (s[2] && s[2].kind !== 'text') return { sec: s[2], abs: false };
      if (c0 && (!s[1] || s[1].kind === 'text')) return { sec: GENERAL_SEC, abs: false };
      return { sec: null };
    }
    if (nonText.length === 1 || s.length === 1) return { sec: s[0], abs: false };
    if (v > 0 || (v === 0 && s.length < 3)) return { sec: s[0], abs: false };
    if (v < 0) return { sec: s[1] && s[1].kind !== 'text' ? s[1] : s[0], abs: !!(s[1] && s[1].kind !== 'text') };
    return { sec: s[2] && s[2].kind !== 'text' ? s[2] : s[0], abs: false };
  }

  /* ------------------------------------------------------------ rendering */
  /** segments: [{s}] text, {pad:ch} width of ch, {fill:ch} repeat to fill */
  function out() { const a = []; a.text = ''; return a; }
  function push(o, s) { if (!s) return; const l = o[o.length - 1]; if (l && l.s != null) l.s += s; else o.push({ s }); }

  /** format a number with a number section */
  function fmtNumber(sec, v, abs, opts) {
    const o = out();
    const neg = v < 0 && !abs;
    let x = Math.abs(v);
    /* Excel ignores % and thousands scaling in scientific formats (the % sign still prints) */
    if (sec.pct && !(sec.expAt >= 0)) x *= Math.pow(100, sec.pct);
    if (sec.scale && !(sec.expAt >= 0)) x /= Math.pow(1000, sec.scale);
    if (sec.frac) return fmtFraction(sec, x, neg, o);
    const toks = sec.toks;
    let intStr, decStr = '', expVal = 0;
    if (sec.expAt >= 0) {
      /* scientific: the exponent is a multiple of the integer placeholder count when that is > 1 (engineering) */
      const ip = Math.max(1, sec.intD);
      let e = x === 0 ? 0 : Math.floor(Math.log10(x));
      if (sec.intD > 1) e = Math.floor(e / ip) * ip;
      else if (sec.intD === 0) e = e + 1;
      let m = x === 0 ? 0 : x / Math.pow(10, e);
      let f = fixed(m, sec.decD);
      /* rounding may carry: 9.99 → 10.0 */
      const intDigits = f.split('.')[0].replace(/^0$/, '').length;
      if (x !== 0 && intDigits > Math.max(1, sec.intD === 0 ? 0 : sec.intD === 1 ? 1 : ip)) { e += sec.intD > 1 ? ip : 1; m = x / Math.pow(10, e); f = fixed(m, sec.decD); }
      [intStr, decStr] = f.split('.');
      decStr = decStr || '';
      expVal = e;
    } else {
      const f = fixed(x, sec.decD);
      [intStr, decStr] = f.split('.');
      decStr = decStr || '';
    }
    /* zero in a scientific format fills every integer placeholder with 0 (Excel: 0 in ##0.0E+0 shows 000.0E+0) */
    if (sec.expAt >= 0 && x === 0) intStr = '0'.repeat(Math.max(1, sec.intD));
    else if (intStr === '0') intStr = '';
    /* sign goes first (before any literal); Excel keeps it even when the shown value rounds to zero */
    if (neg) push(o, '-');
    /* integer placeholders */
    const intToks = toks.slice(0, sec.intEnd);
    const nIntD = sec.intD;
    let digits = intStr;
    /* how many of the placeholders are mandatory zeros etc. */
    const placeholders = intToks.filter((t) => t.t === 'd');
    /* digits aligned right; extra digits go before the first placeholder */
    let di = digits.length - placeholders.length; /* index into digits for the first placeholder */
    let pi = 0;
    const groupDigits = (s, startPos) => {
      if (!sec.group) return s;
      /* insert commas according to position from the right of the whole integer */
      let r = '';
      for (let k = 0; k < s.length; k++) {
        const posFromRight = digits.length - (startPos + k) - 1;
        r += s[k];
        if (posFromRight > 0 && posFromRight % 3 === 0) r += ',';
      }
      return r;
    };
    let emitted = false; /* any integer digit emitted so far */
    for (let k = 0; k < intToks.length; k++) {
      const t = intToks[k];
      if (t.t === 'd') {
        let s = '';
        if (pi === 0 && di > 0) { s = groupDigits(digits.slice(0, di), 0); }
        const idx = di + pi; /* digit index for this placeholder */
        if (idx >= 0 && idx < digits.length) { s += digits[idx]; emitted = true; }
        else if (t.ch === '0') { s += '0'; emitted = true; }
        else if (t.ch === '?') {
          s += ' ';
          /* a separator between two padded positions becomes a space too */
          const posFromRight = nIntD - pi - 1;
          if (sec.group && posFromRight > 0 && posFromRight % 3 === 0 && idx < 0) s += ' ';
        }
        if (sec.group && idx >= 0 && idx < digits.length) {
          const posFromRight = digits.length - idx - 1;
          if (posFromRight > 0 && posFromRight % 3 === 0) s += ',';
        } else if (sec.group && t.ch === '0' && emitted && idx < 0) {
          /* zero-padding placeholders also get separators: 0,000 */
          const posFromRight = nIntD - pi - 1;
          if (posFromRight > 0 && posFromRight % 3 === 0 && posFromRight < nIntD) s += ',';
        }
        push(o, s);
        pi++;
      } else if (t.t === ',') { /* grouping flag */ }
      else emitTok(o, t, neg);
    }
    void emitted;
    /* decimal part */
    if (sec.decAt >= 0) {
      const end = sec.expAt >= 0 ? sec.expAt : toks.length;
      const decToks = toks.slice(sec.decAt + 1, end);
      /* trailing '#' / '?' trim zeros */
      const ph = decToks.filter((t) => t.t === 'd');
      let ds = decStr.split('');
      let lastSig = ds.length - 1;
      while (lastSig >= 0 && ds[lastSig] === '0' && ph[lastSig] && ph[lastSig].ch !== '0') lastSig--;
      push(o, '.');
      let q = 0;
      for (const t of decToks) {
        if (t.t === 'd') {
          if (q <= lastSig) push(o, ds[q]); else if (t.ch === '?') push(o, ' ');
          q++;
        } else emitTok(o, t, neg);
      }
    }
    if (sec.expAt >= 0) {
      const t = toks[sec.expAt];
      const eToks = toks.slice(sec.expAt + 1);
      const ph = eToks.filter((x) => x.t === 'd');
      push(o, t.upper ? 'E' : 'e');
      /* the exponent's sign sits next to its digits, after any literal written between E+ and them */
      const sign = expVal < 0 ? '-' : t.sign === '+' ? '+' : '';
      let es = String(Math.abs(expVal));
      /* missing leading exponent digits: 0 → '0', ? → space, # → nothing */
      if (es.length < ph.length) { let pad = ''; for (let k = 0; k < ph.length - es.length; k++) pad += ph[k].ch === '0' ? '0' : ph[k].ch === '?' ? ' ' : ''; es = pad + es; }
      let first = true;
      if (!ph.length) push(o, sign);
      for (const tk of eToks) {
        if (tk.t === 'd') { if (first) { push(o, sign + es); first = false; } }
        else emitTok(o, tk, neg);
      }
    }
    return o;
  }
  function emitTok(o, t, neg) {
    switch (t.t) {
      case 'lit': push(o, t.s); break;
      case 'pad': o.push({ pad: t.ch }); break;
      case 'fill': o.push({ fill: t.ch }); break;
      case '%': push(o, '%'); break;
      case '.': push(o, '.'); break;
      case '/': push(o, '/'); break;
      case 'gen': break;
      case '@': break;
      case 'skip': break;
      case ',': push(o, ','); break;
      default: break;
    }
    void neg;
  }
  /* best rational approximation with denominator ≤ maxDen (Stern–Brocot / continued fractions) */
  function approx(x, maxDen) {
    let a = Math.floor(x), h1 = 1, h0 = 0, k1 = 0, k0 = 1;
    let b = x;
    let hb = a, kb = 1;
    for (let i = 0; i < 64; i++) {
      a = Math.floor(b);
      const h2 = a * h1 + h0, k2 = a * k1 + k0;
      if (k2 > maxDen) {
        /* semiconvergent check */
        const t = Math.floor((maxDen - k0) / k1);
        const hs = t * h1 + h0, ks = t * k1 + k0;
        if (ks > 0 && Math.abs(x - hs / ks) < Math.abs(x - hb / kb)) { hb = hs; kb = ks; }
        break;
      }
      h0 = h1; h1 = h2; k0 = k1; k1 = k2;
      hb = h2; kb = k2;
      const f = b - a;
      if (f < 1e-12) break;
      b = 1 / f;
    }
    return [hb, kb];
  }
  function fmtFraction(sec, x, neg, o) {
    const dDigits = sec.fracDen.length;
    let ip = 0, num, den;
    if (sec.fracInt) { ip = Math.floor(x); x -= ip; }
    if (sec.fixedDen) { den = sec.fixedDen; num = Math.round(x * den); }
    else {
      const maxDen = Math.pow(10, Math.max(1, dDigits)) - 1;
      [num, den] = approx(x, maxDen);
    }
    if (sec.fracInt && num === den && den > 0) { ip += 1; num = 0; }
    if (neg && (ip || num)) push(o, '-');
    const padTo = (s, toks, left) => {
      /* place digits into placeholders (right aligned for numerator, left for denominator) */
      const n = toks.length;
      let str = s;
      if (str.length < n) {
        const extra = n - str.length;
        let pad = '';
        for (let k = 0; k < extra; k++) {
          const t = left ? toks[str.length + k] : toks[k];
          pad += t && t.ch === '0' ? '0' : t && t.ch === '?' ? ' ' : '';
        }
        str = left ? str + pad : pad + str;
      }
      return str;
    };
    if (sec.fracInt) {
      /* integer part tokens */
      const iToks = sec.fracIntToks;
      const ph = iToks.filter((t) => t.t === 'd');
      let is = ip ? String(ip) : '';
      if (!is && ph.some((t) => t.ch === '0')) is = '0';
      if (!ip && !num && !is) is = '0';
      let placed = false;
      for (const t of iToks) {
        if (t.t === 'd') { if (!placed) { push(o, padTo(is, ph)); placed = true; } }
        else if (t.t === 'lit' && !ip && num && !/[^\s]/.test(t.s) && !is) push(o, t.s);
        else emitTok(o, t);
      }
    } else for (const t of sec.fracIntToks || []) if (t.t !== 'd') emitTok(o, t); /* literals before the numerator */
    if (num === 0 && sec.fracInt) {
      /* whole number: the fraction part becomes spaces */
      const blank = ' '.repeat(sec.fracNum.length + 1 + (sec.fixedDen ? String(sec.fixedDen).length : dDigits));
      for (const t of sec.fracMid) emitTok(o, t);
      push(o, blank);
      for (const t of sec.fracTail) emitTok(o, t);
      return o;
    }
    push(o, padTo(String(num), sec.fracNum));
    for (const t of sec.fracMid) emitTok(o, t);
    push(o, '/');
    for (const t of sec.fracPre || []) emitTok(o, t);
    push(o, sec.fixedDen ? String(den) : padTo(String(den), sec.fracDen, true));
    for (const t of sec.fracTail) emitTok(o, t);
    return o;
  }
  function fmtDate(sec, v, opts) {
    const o = out();
    if (v < 0 && sec.elapsed && v > -2958466) { const r = fmtDate(sec, -v, opts); if (r) r.unshift({ s: '-' }); return r; }
    if (v < 0 || v >= 2958466) return null; /* #### */
    const subTok = sec.toks.find((t) => t.t === 'sub');
    const subDigits = subTok ? subTok.len : 0;
    const p = NF.serialToParts(v, opts && opts.date1904, subDigits);
    const hasAmPm = sec.toks.some((t) => t.t === 'ampm');
    for (const t of sec.toks) {
      switch (t.t) {
        case 'dt': {
          const s = t.s;
          if (t.min) { push(o, s === 'mm' ? String(p.M).padStart(2, '0') : String(p.M)); break; }
          switch (s) {
            case 'yy': case 'y': push(o, String(p.y % 100).padStart(2, '0')); break;
            case 'yyy': case 'yyyy': push(o, String(p.y).padStart(4, '0')); break;
            case 'bb': push(o, String((p.y + 543) % 100).padStart(2, '0')); break;
            case 'bbbb': push(o, String(p.y + 543)); break;
            case 'm': push(o, String(p.m)); break;
            case 'mm': push(o, String(p.m).padStart(2, '0')); break;
            case 'mmm': push(o, MONTHS[p.m - 1].slice(0, 3)); break;
            case 'mmmm': push(o, MONTHS[p.m - 1]); break;
            case 'mmmmm': push(o, MONTHS[p.m - 1][0]); break;
            case 'd': push(o, String(p.d)); break;
            case 'dd': push(o, String(p.d).padStart(2, '0')); break;
            case 'ddd': push(o, DAYS[p.wd].slice(0, 3)); break;
            case 'dddd': push(o, DAYS[p.wd]); break;
            case 'h': push(o, String(hasAmPm ? ((p.H + 11) % 12) + 1 : p.H)); break;
            case 'hh': push(o, String(hasAmPm ? ((p.H + 11) % 12) + 1 : p.H).padStart(2, '0')); break;
            case 's': push(o, String(p.S)); break;
            case 'ss': push(o, String(p.S).padStart(2, '0')); break;
            default: push(o, s);
          }
          break;
        }
        case 'el': {
          const total = p.days * 86400 + p.secs;
          let val;
          if (t.u === 'h') val = Math.floor(total / 3600);
          else if (t.u === 'm') val = Math.floor(total / 60);
          else val = Math.floor(total);
          push(o, String(val).padStart(t.len, '0'));
          break;
        }
        case 'sub': push(o, '.' + fixed(p.sub, t.len).split('.')[1]); break;
        case 'ampm': {
          const pm = p.H >= 12;
          if (/^a\/p$/i.test(t.s)) { const c = pm ? t.s[2] : t.s[0]; push(o, c); }
          else push(o, pm ? 'PM' : 'AM');
          break;
        }
        case 'd': push(o, t.ch === '0' ? '0' : ''); break;
        default: emitTok(o, t);
      }
    }
    return o;
  }
  function fmtText(sec, s) {
    const o = out();
    for (const t of sec.toks) {
      if (t.t === '@') push(o, s);
      else if (t.t === 'gen') push(o, s);
      else if (t.t === 'd' || t.t === 'dt' || t.t === 'el' || t.t === 'ampm' || t.t === 'sub' || t.t === 'exp') { /* number placeholders in a text section print nothing */ }
      else emitTok(o, t);
    }
    return o;
  }
  const segText = (segs) => segs.map((x) => (x.s != null ? x.s : x.pad != null ? ' ' : '')).join('');

  /**
   * Format a value. value: number | string | boolean | {e:'#N/A'} | null
   * opts: {date1904, width (chars available for General)}
   * returns {text, segs, color, hashes(bool when the number cannot be shown)}
   */
  NF.format = function (code, value, opts) {
    const p = NF.parse(code);
    let segs = null, sec = null;
    if (value == null || value === '') return { text: '', segs: [], color: null };
    if (typeof value === 'boolean') value = value ? 'TRUE' : 'FALSE';
    else if (typeof value === 'object' && value.e) return { text: value.e, segs: [{ s: value.e }], color: null };
    if (typeof value === 'string') {
      const s = p.secs;
      const ts = s.length >= 4 ? s[3] : s.find((x) => x.kind === 'text') || null;
      if (ts) { sec = ts; segs = fmtText(ts, value); }
      else return { text: value, segs: [{ s: value }], color: null };
      return { text: segText(segs), segs, color: sec.color };
    }
    if (typeof value !== 'number' || !isFinite(value)) return { text: String(value), segs: [{ s: String(value) }], color: null };
    const pick = pickNum(p, value);
    sec = pick.sec;
    if (!sec) return { text: '#'.repeat(10), segs: [{ s: '##########' }], color: null, hashes: true };
    if (sec.kind === 'general' || (sec.kind === 'text' && p.secs.length === 1) || sec.kind === 'lit' && !sec.toks.length) {
      const g = NF.general(pick.abs ? Math.abs(value) : value, opts && opts.width, opts && opts.fn);
      if (sec.kind === 'general') { segs = out(); for (const t of sec.toks) { if (t.t === 'gen') push(segs, g); else emitTok(segs, t); } }
      else segs = [{ s: g }];
    } else if (sec.kind === 'date') {
      segs = fmtDate(sec, value, opts);
      if (!segs) return { text: '#'.repeat(10), segs: [{ s: '##########' }], color: sec.color, hashes: true };
    } else if (sec.kind === 'num') segs = fmtNumber(sec, value, pick.abs, opts);
    else if (sec.kind === 'text') { segs = out(); for (const t of sec.toks) { if (t.t !== '@') emitTok(segs, t); else push(segs, NF.general(value, opts && opts.width)); } }
    else { segs = out(); for (const t of sec.toks) emitTok(segs, t); if (value < 0 && !pick.abs && sec === p.secs[0]) segs.unshift({ s: '-' }); }
    return { text: segText(segs), segs, color: sec.color };
  };
  /** plain text (TEXT function, CSV "as displayed") */
  NF.text = (code, value, opts) => NF.format(code, value, opts).text;

  /* ------------------------------------------------------- classification */
  NF.isDate = (code) => { const p = NF.parse(code); return p.secs.length > 0 && p.secs[0].kind === 'date'; };
  NF.isTimeOnly = (code) => { const p = NF.parse(code); const s = p.secs[0]; return !!s && s.kind === 'date' && !s.toks.some((t) => t.t === 'dt' && !t.min && /^[ymdbe]/.test(t.s)); };
  NF.isPercent = (code) => { const p = NF.parse(code); return !!(p.secs[0] && p.secs[0].pct); };
  /** category for the Format Cells dialog */
  NF.category = function (code) {
    if (!code || code === 'General') return 'General';
    const p = NF.parse(code);
    const s0 = p.secs[0];
    if (!s0) return 'General';
    if (s0.kind === 'text' && p.secs.length === 1) return 'Text';
    if (s0.kind === 'date') return NF.isTimeOnly(code) ? 'Time' : 'Date';
    if (s0.kind !== 'num') return 'Custom';
    if (s0.frac) return 'Fraction';
    if (s0.expAt >= 0) return 'Scientific';
    if (s0.pct) return 'Percentage';
    if (/^_\(\S*\*/.test(code) || /\* /.test(code)) return 'Accounting';
    if (/\$|€|£|¥|\[\$/.test(code)) return 'Currency';
    if (/^0(\.0+)?$|^#,##0(\.0+)?$/.test(s0.src) || /^#,##0(\.0+)?_\)/.test(code) || /^0(\.0+)?_\)/.test(code)) return 'Number';
    return 'Custom';
  };
  /** decimals shown by a number format (for Increase/Decrease Decimal) */
  NF.decimals = (code) => { const p = NF.parse(code); const s = p.secs[0]; return s && s.kind === 'num' && !s.frac ? s.decD : 0; };
  /** change the number of decimals in every number section of a code */
  NF.withDecimals = function (code, delta, value) {
    if (!code || code === 'General') {
      /* General → as many decimals as the value shows, ±1 */
      let cur = 0;
      if (typeof value === 'number') { const g = NF.general(value); const m = /\.(\d+)/.exec(g); cur = m && !/E/.test(g) ? m[1].length : 0; }
      const n = Math.max(0, cur + delta);
      return n ? '0.' + '0'.repeat(n) : '0';
    }
    return splitSections(code).map((sec) => {
      if (/^\s*$/.test(sec) || NF.parse(sec).secs[0].kind !== 'num') return sec;
      /* find the last digit placeholder outside quotes/brackets */
      let q = false, b = false, lastDigit = -1, dot = -1;
      for (let i = 0; i < sec.length; i++) {
        const ch = sec[i];
        if (q) { if (ch === '"') q = false; continue; }
        if (ch === '"') { q = true; continue; }
        if (ch === '\\' || ch === '_' || ch === '*') { i++; continue; }
        if (ch === '[') { b = true; continue; }
        if (ch === ']') { b = false; continue; }
        if (b) continue;
        if (ch === 'E' || ch === 'e') break;
        if (ch === '0' || ch === '#' || ch === '?') lastDigit = i;
        if (ch === '.' && dot < 0) dot = i;
      }
      if (lastDigit < 0) return sec;
      if (delta > 0) return dot >= 0 ? sec.slice(0, lastDigit + 1) + '0'.repeat(delta) + sec.slice(lastDigit + 1) : sec.slice(0, lastDigit + 1) + '.' + '0'.repeat(delta) + sec.slice(lastDigit + 1);
      if (dot < 0) return sec;
      const decs = lastDigit - dot;
      const remove = Math.min(decs, -delta);
      if (remove === decs) return sec.slice(0, dot) + sec.slice(lastDigit + 1);
      return sec.slice(0, lastDigit + 1 - remove) + sec.slice(lastDigit + 1);
    }).join(';');
  };

  /* ------------------------------------------------------- input parsing */
  const MON3 = MONTHS.map((m) => m.slice(0, 3).toLowerCase());
  const monthIndex = (s) => { const k = s.toLowerCase(); let i = MONTHS.findIndex((m) => m.toLowerCase() === k); if (i < 0) i = MON3.indexOf(k.slice(0, 3)); if (i < 0 && k === 'sept') i = 8; return i < 0 ? -1 : (k.length >= 3 && (MONTHS[i].toLowerCase().startsWith(k) || k === 'sept') ? i : -1); };
  function timeFrom(str) {
    const m = /^(\d{1,2}):(\d{1,2})(?::(\d{1,2}(?:\.\d+)?))?\s*(am|pm|a|p)?$/i.exec(str) || /^(\d{1,2})\s*(am|pm|a|p)$/i.exec(str);
    if (!m) return null;
    let H, M = 0, S = 0, ap;
    if (m.length === 3) { H = +m[1]; ap = m[2]; } else { H = +m[1]; M = +m[2]; S = m[3] ? +m[3] : 0; ap = m[4]; }
    if (ap) { if (H > 12 || H === 0 && false) return null; const pm = /^p/i.test(ap); if (H === 12) H = pm ? 12 : 0; else if (pm) H += 12; }
    if (M > 59 || S >= 60) return null;
    const fmt = ap ? (m[3] ? 'h:mm:ss AM/PM' : 'h:mm AM/PM') : m[3] ? 'h:mm:ss' : H >= 24 ? '[h]:mm' : 'h:mm';
    return { v: (H * 3600 + M * 60 + S) / 86400, fmt };
  }
  const validYMD = (y, m, d) => m >= 1 && m <= 12 && d >= 1 && d <= new Date(Date.UTC(y, m, 0)).getUTCDate() + (y === 1900 && m === 2 ? 1 : 0);
  const fullYear = (y, s) => (s.length <= 2 ? (y < 30 ? 2000 + y : 1900 + y) : y);
  function dateFrom(str, date1904) {
    const now = new Date();
    let m;
    const mk = (y, mo, d, fmt) => (validYMD(y, mo, d) && y >= (date1904 ? 1904 : 1900) && y <= 9999 ? { v: NF.dateToSerial(y, mo, d, date1904), fmt } : null);
    if ((m = /^(\d{1,2})[/-](\d{1,2})[/-](\d{1,4})$/.exec(str))) return mk(fullYear(+m[3], m[3]), +m[1], +m[2], 'm/d/yyyy');
    if ((m = /^(\d{4})[/-](\d{1,2})[/-](\d{1,2})$/.exec(str))) return mk(+m[1], +m[2], +m[3], 'm/d/yyyy');
    if ((m = /^(\d{1,2})[/-](\d{1,2})$/.exec(str))) { const r = mk(now.getFullYear(), +m[1], +m[2], 'd-mmm'); if (r) return r; return +m[2] > 31 || true ? (validYMD(fullYear(+m[2], m[2]), +m[1], 1) ? mk(fullYear(+m[2], m[2]), +m[1], 1, 'mmm-yy') : null) : null; }
    if ((m = /^(\d{1,2})[\s-]([a-z]{3,9})\.?(?:[\s-,]+(\d{2,4}))?$/i.exec(str))) { const mo = monthIndex(m[2]); if (mo < 0) return null; return m[3] ? mk(fullYear(+m[3], m[3]), mo + 1, +m[1], 'd-mmm-yy') : mk(now.getFullYear(), mo + 1, +m[1], 'd-mmm'); }
    if ((m = /^([a-z]{3,9})\.?[\s-](\d{1,2})(?:(?:,\s*|\s+|-)(\d{2,4}))?$/i.exec(str))) {
      const mo = monthIndex(m[1]); if (mo < 0) return null;
      if (m[3]) return mk(fullYear(+m[3], m[3]), mo + 1, +m[2], 'd-mmm-yy');
      /* "Jan 24": a day when it can be one, otherwise a year */
      const r = +m[2] <= 31 ? mk(now.getFullYear(), mo + 1, +m[2], 'd-mmm') : null;
      return r || mk(fullYear(+m[2], m[2]), mo + 1, 1, 'mmm-yy');
    }
    if ((m = /^([a-z]{3,9})\.?[\s-](\d{4})$/i.exec(str))) { const mo = monthIndex(m[1]); if (mo < 0) return null; return mk(+m[2], mo + 1, 1, 'mmm-yy'); }
    return null;
  }
  /**
   * Interpret typed input the way Excel does (en-US).
   * returns {v, fmt?} — v: number | string | boolean | {e} ; fmt: number format to apply when the cell is General
   */
  NF.parseInput = function (text, opts) {
    const date1904 = opts && opts.date1904;
    if (text == null) return { v: null };
    const raw = String(text);
    if (raw === '') return { v: null };
    const s = raw.trim();
    if (!s) return { v: raw };
    const up = s.toUpperCase();
    if (up === 'TRUE') return { v: true };
    if (up === 'FALSE') return { v: false };
    if (/^#(NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|GETTING_DATA|SPILL!|CALC!)$/.test(up)) return { v: { e: up } };
    const n = NF.parseNumber(s);
    /* beyond ±9.99999999999999E+307 Excel keeps the entry as text */
    if (n) return typeof n.v === 'number' && !isFinite(n.v) ? { v: raw } : n;
    /* date + time */
    let m = /^(.+?)\s+(\d{1,2}:\d{1,2}(?::\d{1,2}(?:\.\d+)?)?(?:\s*[ap]m?)?)$/i.exec(s);
    if (m) {
      const d = dateFrom(m[1], date1904), t = timeFrom(m[2]);
      if (d && t) return { v: d.v + t.v, fmt: /[ap]/i.test(m[2]) ? 'm/d/yyyy h:mm AM/PM' : 'm/d/yyyy h:mm' };
    }
    const t = timeFrom(s);
    if (t) return t;
    const d = dateFrom(s, date1904);
    if (d) return d;
    return { v: raw };
  };
  /** numbers: 1,234.5  $1,234  (12)  -1.5e3  45%  0 1/2  1 3/4 */
  NF.parseNumber = function (s) {
    s = String(s).trim();
    let m;
    /* fractions "a b/c" */
    if ((m = /^([+-])?(\d+)\s+(\d+)\/(\d+)$/.exec(s))) {
      const den = +m[4]; if (!den || +m[3] >= den) return null;
      const v = (+m[2] + +m[3] / den) * (m[1] === '-' ? -1 : 1);
      return { v, fmt: den < 10 ? '# ?/?' : '# ??/??' };
    }
    let neg = false, cur = false, pct = false, paren = false;
    let x = s;
    if (/^\(.*\)$/.test(x)) { paren = true; x = x.slice(1, -1).trim(); }
    if (/^[+-]/.test(x)) { neg = x[0] === '-'; x = x.slice(1).trim(); }
    if (/^\$/.test(x)) { cur = true; x = x.slice(1).trim(); if (/^[+-]/.test(x)) { neg = x[0] === '-'; x = x.slice(1); } }
    if (/\$$/.test(x)) { cur = true; x = x.slice(0, -1).trim(); }
    if (/%$/.test(x)) { pct = true; x = x.slice(0, -1).trim(); }
    if (/^[+-]/.test(x) && !neg) { neg = x[0] === '-'; x = x.slice(1); }
    if (!/^(\d{1,3}(,\d{3})+|\d*)(\.\d*)?([eE][+-]?\d+)?$/.test(x) || !/\d/.test(x)) return null;
    const grouped = /,/.test(x);
    const sci = /e/i.test(x);
    if (sci && (pct || cur || grouped)) return null;
    let v = parseFloat(x.replace(/,/g, ''));
    if (isNaN(v)) return null;
    if (paren) { if (neg) return null; neg = true; }
    if (neg) v = -v;
    const decs = (/\.(\d+)/.exec(x) || [, ''])[1].length;
    let fmt = null;
    if (pct) { v /= 100; fmt = decs ? '0.' + '0'.repeat(decs) + '%' : '0%'; }
    else if (cur) fmt = decs ? '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)' : '"$"#,##0_);[Red]\\("$"#,##0\\)';
    else if (sci) fmt = '0.00E+00';
    else if (grouped) fmt = decs ? '#,##0.00' : '#,##0';
    else if (paren) fmt = null;
    return { v, fmt };
  };
})(typeof window !== 'undefined' ? window : globalThis);
