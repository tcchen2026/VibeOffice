/* Quire — built-in document templates (File ▸ New ▸ On my computer…). Placeholders are MACROBUTTON
 * fields: clicking one selects it so that typing replaces it, as in Word's templates.
 */
(function () {
  'use strict';
  const L = window.L, D = L.D;
  const T = (L.templates = {});
  let fidN = 0;
  const fid = () => 'tf' + (++fidN) + '_' + D.nid();
  const P = (runs, pPr, rPr) => D.para(runs, pPr || {}, rPr || {});
  const t = (s, r) => D.text(s, r || {});
  /** [Click here and type …] placeholder */
  const ph = (label, r) => { const f = fid(); return [D.item('fb', { fid: f, instr: ' MACROBUTTON  DoFieldClick [' + label + '] ' }, r || {}), D.item('fs', { fid: f }, r || {}), D.text('[' + label + ']', r || {}), D.item('fe', { fid: f }, r || {})]; };
  const field = (instr, result, r) => { const f = fid(); return [D.item('fb', { fid: f, instr: ' ' + instr + ' ' }, r || {}), D.item('fs', { fid: f }, r || {}), D.text(result, r || {}), D.item('fe', { fid: f }, r || {})]; };
  const today = () => L.fields ? L.fields.formatDate(new Date(), 'MMMM d, yyyy') : new Date().toDateString();
  const style = (d, id, name, o) => { d.styles[id] = Object.assign({ id, name, type: 'paragraph', basedOn: 'Normal', next: null, pPr: {}, rPr: {}, custom: true, q: true }, o); };
  const LIST = [
    { id: 'letter', name: 'Contemporary Letter', cat: 'Letters & Faxes' },
    { id: 'memo', name: 'Professional Memo', cat: 'Memos' },
    { id: 'fax', name: 'Elegant Fax', cat: 'Letters & Faxes' },
    { id: 'resume', name: 'Contemporary Résumé', cat: 'Other Documents' },
    { id: 'report', name: 'Professional Report', cat: 'Reports' },
    { id: 'agenda', name: 'Meeting Agenda', cat: 'Other Documents' },
  ];
  T.list = () => LIST;
  T.categories = () => ['General', 'Letters & Faxes', 'Memos', 'Reports', 'Other Documents'];
  const B = {
    letter(d) {
      d.sect.mt = 72; d.sect.ml = d.sect.mr = 90;
      style(d, 'CompanyName', 'Company Name', { pPr: { sp: { a: 24 }, borders: { bottom: { val: 'single', sz: 0.75, color: '808080', space: 4 } } }, rPr: { font: 'Arial Black', sz: 16, spacing: -0.5 } });
      style(d, 'ReturnAddress', 'Return Address', { pPr: { jc: 'right' }, rPr: { font: 'Arial', sz: 8 } });
      style(d, 'InsideAddress', 'Inside Address', { rPr: { font: 'Arial', sz: 10 } });
      style(d, 'BodyTextLetter', 'Body Text Letter', { pPr: { sp: { a: 12, line: 1.15, rule: 'auto' }, jc: 'both' }, rPr: { font: 'Arial', sz: 10 } });
      d.main.blocks = [
        P(ph('Type return address'), { style: 'ReturnAddress' }),
        P(ph('Company Name Here', { font: 'Arial Black' }), { style: 'CompanyName' }),
        P(field('DATE \\@ "MMMM d, yyyy"', today()), { style: 'InsideAddress', sp: { a: 24 } }),
        P(ph('Type recipient’s address'), { style: 'InsideAddress', sp: { a: 24 } }),
        P([t('Dear '), ...ph('Recipient'), t(':')], { style: 'BodyTextLetter' }),
        P([t('Write the body of your letter here. To change the look of every paragraph at once, open Format ▸ Styles and Formatting and modify the style instead of formatting text by hand.')], { style: 'BodyTextLetter' }),
        P([t('To return to the main view, simply keep typing. Placeholders like the ones above disappear as soon as you click them and type.')], { style: 'BodyTextLetter' }),
        P([t('Sincerely,')], { style: 'BodyTextLetter', sp: { a: 36 } }),
        P(ph('Type your name'), { style: 'InsideAddress' }),
        P(ph('Type job title'), { style: 'InsideAddress' }),
      ];
    },
    memo(d) {
      style(d, 'MemoTitle', 'Memo Title', { pPr: { sp: { a: 24 }, borders: { bottom: { val: 'single', sz: 3, color: '000000', space: 6 } } }, rPr: { font: 'Arial Black', sz: 36, spacing: -1.5 } });
      style(d, 'MessageHeader', 'Message Header', { pPr: { sp: { a: 6 }, ind: { l: 1080, fl: -1080 }, tabs: [{ pos: 72, al: 'left' }] }, rPr: { font: 'Arial', sz: 10 } });
      style(d, 'MemoBody', 'Memo Body', { pPr: { sp: { a: 12 } }, rPr: { font: 'Arial', sz: 10 } });
      d.styles.MessageHeader.pPr.ind = { l: 72, fl: -72 };
      const label = (s) => t(s, { b: true, caps: true });
      d.main.blocks = [
        P(ph('Company Name Here'), { style: 'MemoBody', jc: 'right' }),
        P([t('Memo')], { style: 'MemoTitle' }),
        P([label('To:'), D.item('tab'), ...ph('Type name')], { style: 'MessageHeader' }),
        P([label('From:'), D.item('tab'), ...ph('Type name')], { style: 'MessageHeader' }),
        P([label('CC:'), D.item('tab'), ...ph('Type name')], { style: 'MessageHeader' }),
        P([label('Date:'), D.item('tab'), ...field('DATE \\@ "M/d/yyyy"', L.fields ? L.fields.formatDate(new Date(), 'M/d/yyyy') : '')], { style: 'MessageHeader' }),
        P([label('Re:'), D.item('tab'), ...ph('Type subject')], { style: 'MessageHeader', borders: { bottom: { val: 'single', sz: 0.75, color: '000000', space: 8 } }, sp: { a: 18 } }),
        P([t('How to Use This Memo Template', { b: true })], { style: 'MemoBody' }),
        P([t('Start the memo with the point you need the reader to act on. Keep the headings short; the Style box on the Formatting toolbar has Heading 1–3 for structure.')], { style: 'MemoBody' }),
        P([t('The decorative shapes are ordinary drawing objects: click one and press Delete to remove it, or drag it somewhere else.')], { style: 'MemoBody' }),
      ];
    },
    fax(d) {
      style(d, 'FaxTitle', 'Fax Title', { pPr: { jc: 'center', sp: { b: 24, a: 24 } }, rPr: { font: 'Garamond', sz: 36, caps: true, spacing: 6 } });
      style(d, 'FaxLabel', 'Fax Label', { rPr: { font: 'Garamond', sz: 10, caps: true, b: true } });
      d.main.blocks = [P(ph('Company Name'), { jc: 'center' }, { font: 'Garamond', sz: 14, smallCaps: true }), P([t('Facsimile Transmittal')], { style: 'FaxTitle' })];
      const tb = D.simpleTable(4, 4, 432, { tblPr: { style: 'TableNormal' } });
      const cells = [['To:', 'Click here and type name', 'From:', 'Click here and type name'], ['Fax:', 'Click here and type fax number', 'Pages:', 'Click here and type number of pages'], ['Phone:', 'Click here and type phone number', 'Date:', null], ['Re:', 'Click here and type subject of fax', 'CC:', 'Click here and type name']];
      cells.forEach((r, ri) => r.forEach((v, ci) => { const p = tb.rows[ri].cells[ci].blocks[0]; if (ci % 2 === 0) { p.runs = [t(v)]; p.pPr.style = 'FaxLabel'; } else p.runs = v ? ph(v, { font: 'Garamond' }) : field('DATE \\@ "MMMM d, yyyy"', today(), { font: 'Garamond' }); }));
      tb.grid = [60, 156, 60, 156];
      L.tables && L.tables.fixWidths(tb);
      d.main.blocks.push(tb, P([t('☐ Urgent   ☐ For Review   ☐ Please Comment   ☐ Please Reply   ☐ Please Recycle', { font: 'Garamond', sz: 10 })], { sp: { b: 18, a: 18 }, borders: { top: { val: 'single', sz: 0.75, color: '000000', space: 6 }, bottom: { val: 'single', sz: 0.75, color: '000000', space: 6 } } }),
        P([t('Comments: ', { b: true, caps: true, font: 'Garamond' }), ...ph('Type any comments', { font: 'Garamond' })], { sp: { b: 12 } }));
    },
    resume(d) {
      d.sect.ml = d.sect.mr = 90;
      style(d, 'Name', 'Name', { pPr: { sp: { a: 6 }, borders: { bottom: { val: 'single', sz: 1.5, color: '000000', space: 6 } } }, rPr: { font: 'Arial', sz: 22, b: true, spacing: -1 } });
      style(d, 'Address1', 'Address 1', { pPr: { jc: 'right' }, rPr: { font: 'Arial', sz: 8 } });
      style(d, 'SectionTitle', 'Section Title', { pPr: { sp: { b: 18, a: 6 }, keepNext: true }, rPr: { font: 'Arial', sz: 10, b: true, caps: true, spacing: 1 } });
      style(d, 'Objective', 'Objective', { pPr: { sp: { a: 12 } }, rPr: { font: 'Arial', sz: 10 } });
      style(d, 'CompanyNameRes', 'Company Name (Résumé)', { pPr: { tabs: [{ pos: 160, al: 'left' }, { pos: 432, al: 'right' }], sp: { b: 6 } }, rPr: { font: 'Arial', sz: 10 } });
      style(d, 'JobTitle', 'Job Title', { pPr: { sp: { a: 4 } }, rPr: { font: 'Arial Black', sz: 9 } });
      style(d, 'Achievement', 'Achievement', { pPr: { ind: { l: 18, fl: -18 }, sp: { a: 3 } }, rPr: { font: 'Arial', sz: 10 } });
      const absId = D.addNum(d, D.makeBulletAbs('•', 'Symbol', ''));
      const ach = (s) => P([t(s)], { style: 'Achievement', num: { id: absId, lvl: 0 } });
      d.main.blocks = [
        P(ph('Street Address • City, ST ZIP Code • Phone • E-mail'), { style: 'Address1' }),
        P(ph('Your Name'), { style: 'Name' }),
        P([t('Objective')], { style: 'SectionTitle' }), P(ph('Type your objective here'), { style: 'Objective' }),
        P([t('Experience')], { style: 'SectionTitle' }),
        P([...ph('Start – End'), D.item('tab'), ...ph('Company Name', { b: true }), D.item('tab'), ...ph('City, State')], { style: 'CompanyNameRes' }),
        P(ph('Job Title'), { style: 'JobTitle' }),
        ach('Describe a result you achieved, with a number if you can.'), ach('Name a responsibility and what changed because of you.'), ach('Add one line per achievement; keep each to a single sentence.'),
        P([t('Education')], { style: 'SectionTitle' }),
        P([...ph('Start – End'), D.item('tab'), ...ph('School Name', { b: true }), D.item('tab'), ...ph('City, State')], { style: 'CompanyNameRes' }),
        P(ph('Degree, major and honors'), { style: 'Achievement' }),
        P([t('Interests')], { style: 'SectionTitle' }), P(ph('Clubs, volunteer work and interests that support your application'), { style: 'Objective' }),
        P([t('Tips')], { style: 'SectionTitle' }), P([t('Click any bracketed placeholder and type over it. The Section Title, Job Title and Achievement styles in the Style box keep every entry lined up.')], { style: 'Objective' }),
      ];
    },
    report(d) {
      d.styles.Title.rPr = { font: 'Arial Black', sz: 32, b: false, kern: 14 };
      d.styles.Title.pPr = { jc: 'left', sp: { b: 180, a: 12 }, outline: 0, borders: { bottom: { val: 'single', sz: 3, color: '000000', space: 6 } } };
      d.styles.Subtitle.rPr = { font: 'Arial', sz: 14, i: false, color: '595959' };
      d.styles.Subtitle.pPr = { jc: 'left', sp: { a: 120 } };
      d.styles.Heading1.rPr = { font: 'Arial Black', sz: 16, b: false };
      d.styles.Heading1.pPr = Object.assign({}, d.styles.Heading1.pPr, { pageBreakBefore: false, sp: { b: 24, a: 6 } });
      d.styles.Heading2.rPr = { font: 'Arial', sz: 13, b: true };
      d.styles.Normal.rPr = { font: 'Garamond', sz: 11 };
      d.styles.Normal.pPr = { sp: { a: 8, line: 1.15, rule: 'auto' } };
      d.main.blocks = [
        P(ph('Company Name'), { jc: 'left' }, { font: 'Arial', sz: 10, caps: true, b: true }),
        P(ph('Type Report Title Here'), { style: 'Title' }),
        P([t('A Professional Report')], { style: 'Subtitle' }),
        P([...field('DATE \\@ "MMMM d, yyyy"', today())], {}, {}),
        P([t('Table of Contents')], { style: 'TOCHeading', pageBreakBefore: true }),
        P([t('Press F9 in the table of contents below to bring it up to date.', { i: true, color: '7F7F7F' })]),
        P([D.item('fb', { fid: 'toc1', instr: ' TOC \\o "1-3" \\h \\z \\u ' }), D.item('fs', { fid: 'toc1' }), t('Right-click and choose Update Field to build the table of contents.'), D.item('fe', { fid: 'toc1' })]),
        P([t('Introduction')], { style: 'Heading1', pageBreakBefore: true }),
        P([t('Replace this text with your own. Headings set in Heading 1 and Heading 2 feed the table of contents, which rebuilds itself when you update it.')]),
        P([t('Using This Report')], { style: 'Heading2' }),
        P([t('Change the margins and paper size with File ▸ Page Setup. Add a header or footer with View ▸ Header and Footer; page numbers come from Insert ▸ Page Numbers.')]),
        P([t('Findings')], { style: 'Heading1' }),
        P([t('Describe your findings here. Insert a chart with Insert ▸ Picture ▸ Chart, and add a caption with Insert ▸ Reference ▸ Caption.')]),
        P([t('Conclusion')], { style: 'Heading1' }),
        P([t('Summarize your conclusions and recommendations.')]),
      ];
      d.hf.rf = { kind: 'ftr', id: 'rf', blocks: [D.para([D.item('tab'), D.item('tab'), D.text('Page '), ...field('PAGE', '1')], { style: 'Footer' })] };
      d.sect.refs = { hdr: {}, ftr: { default: 'rf' } };
      d.styles.TOCHeading = d.styles.TOCHeading || D.builtinStyles().TOCHeading;
    },
    agenda(d) {
      style(d, 'AgendaTitle', 'Agenda Title', { pPr: { jc: 'center', sp: { a: 6 } }, rPr: { font: 'Tahoma', sz: 26, b: true, color: '1F3864' } });
      style(d, 'AgendaItem', 'Agenda Item', { pPr: { tabs: [{ pos: 360, al: 'left' }], sp: { a: 6 } }, rPr: { font: 'Tahoma', sz: 10 } });
      d.main.blocks = [P(ph('Meeting Title'), { style: 'AgendaTitle' }), P([...field('DATE \\@ "dddd, MMMM d, yyyy"', L.fields ? L.fields.formatDate(new Date(), 'dddd, MMMM d, yyyy') : ''), t('  •  '), ...ph('Time'), t('  •  '), ...ph('Location')], { jc: 'center', sp: { a: 18 } }, { font: 'Tahoma', sz: 10 })];
      const tb = D.simpleTable(6, 3, 432);
      const rows = [['Time', 'Item', 'Owner'], ['9:00 am', 'Welcome and introductions', 'Chair'], ['9:15 am', 'Review of last meeting', 'Secretary'], ['9:30 am', 'Project updates', 'Team leads'], ['10:15 am', 'Open discussion', 'All'], ['10:45 am', 'Action items and next steps', 'Chair']];
      rows.forEach((r, ri) => r.forEach((v, ci) => { const p = tb.rows[ri].cells[ci].blocks[0]; p.runs = [t(v, ri === 0 ? { b: true, color: 'FFFFFF' } : {})]; p.pPr.style = 'AgendaItem'; if (ri === 0) tb.rows[ri].cells[ci].tcPr.shd = { val: 'clear', fill: '1F3864' }; }));
      tb.rows[0].trPr.header = true;
      tb.grid = [72, 252, 108];
      L.tables && L.tables.fixWidths(tb);
      d.main.blocks.push(tb, P([t('Additional Instructions:', { b: true })], { style: 'AgendaItem', sp: { b: 18 } }), P(ph('Type any notes or preparation needed'), { style: 'AgendaItem' }));
    },
  };
  T.build = function (id) {
    const d = D.newDoc();
    const prev = D.doc;
    D.doc = d;
    try { if (B[id]) B[id](d); } finally { D.doc = prev; }
    d.props.title = '';
    d.props.template = (LIST.find((x) => x.id === id) || {}).name || 'Normal';
    return d;
  };
  T.create = function (id) {
    const d = T.build(id);
    L.app.newDocument({ build: (nd) => Object.assign(nd, d) });
  };
  /* clicking a MACROBUTTON placeholder selects the whole field so typing replaces it */
  L.bus.on('sel', () => {
    const E = L.ed;
    if (!E.sel || !E.collapsed() || !L.fields || T.busy) return;
    const f = L.fields.at(E.sel.f);
    if (!f || !/^\s*MACROBUTTON\b/i.test(f.it.instr)) return;
    const g = L.O.findField(f.story, f.fid);
    if (!g.begin || !g.end) return;
    const c = E.sel.f;
    if ((c.p === g.begin.p && c.o <= g.begin.o) || (c.p === g.end.p && c.o >= g.end.o)) return;
    T.busy = true;
    E.setSel(D.pos(g.begin.p, g.begin.o), D.pos(g.end.p, g.end.o), { noScroll: true });
    T.busy = false;
  });
})();
