/* Lectern — content templates: ready-made decks for everyday talks, each on one of the design
 * templates (model.js). Used by the Start Center (lectern/?template=<id>); the text on every slide
 * says what to put there.
 */
(function () {
  'use strict';
  const L = window.L, M = L.model, T = L.txt;
  const LIST = [
    { id: 'status', name: 'Project Status Report', design: 'blueprint' },
    { id: 'meeting', name: 'Team Meeting', design: 'meadow' },
    { id: 'lesson', name: 'Lesson', design: 'parchment' },
    { id: 'event', name: 'Event Announcement', design: 'night' },
  ];
  const TL = (L.templates = {});
  TL.list = () => LIST;

  /** a table in place of the slide's table placeholder; rows[0] is the header */
  function table(pres, s, tp, rows, widths) {
    const t = M.newTable(pres, s, rows.length, rows[0].length, tp.x, tp.y, tp.w);
    t.ph = L.clone(tp.ph);
    const rh = Math.min(44, tp.h / rows.length);
    rows.forEach((row, ri) => row.forEach((v, ci) => {
      const cell = t.tbl.rows[ri].cells[ci];
      cell.tx.ps = [T.para(v, { algn: 'l' }, ri === 0 ? { sz: 18, b: true, color: 'lt1' } : { sz: 16 })];
      cell.fill = ri === 0 ? { t: 'solid', c: 'accent1', a: 1 } : { t: 'solid', c: ri % 2 ? '#F2F5FA' : '#FFFFFF', a: 1 };
      for (const k of ['l', 'r', 't', 'b']) cell.bd[k] = { c: '#B4C3DC', w: 1, dash: 'solid' };
      cell.tx.anchor = 'ctr';
    }));
    t.tbl.rows.forEach((r) => { r.h = rh; });
    t.h = rh * rows.length;
    t.tbl.cols = widths.map((f) => tp.w * f);
    s.shapes.splice(s.shapes.indexOf(tp), 1, t);
    return t;
  }

  /** a SmartArt diagram in place of the slide's body placeholder (js/diagram.js) */
  const diagram = (pres, s, layout, lines, opts) => L.diagram.make(pres, s, s.shapes.find((x) => x.ph && x.ph.type === 'body'), layout, lines, opts);

  const B = {
    status(add, pres) {
      add('title', 'Project Name', ['Status report · Reporting period', 'Presented by your name']);
      add('text', 'Summary', ['Overall status: On track / At risk / Off track', 'One sentence on progress since the last report', 'One sentence on what needs attention']);
      diagram(pres, add('text', 'Project Phases'), 'chevron', ['Discover', 'Design', 'Build', 'Launch']);
      const s = add('table', 'Milestones');
      table(pres, s, s.shapes.find((x) => x.ph && x.ph.type === 'tbl'), [['Milestone', 'Owner', 'Due', 'Status'], ['Requirements signed off', 'Name', 'Date', 'Done'], ['Design complete', 'Name', 'Date', 'On track'], ['First release', 'Name', 'Date', 'At risk'], ['Launch', 'Name', 'Date', 'Not started']], [0.42, 0.2, 0.16, 0.22]);
      add('twoText', 'Done and Next', ['Done this period', '\tWhat was finished', '\tWhat was delivered'], ['Next period', '\tWhat comes next', '\tWho does it']);
      add('text', 'Risks and Issues', ['Risk: what could go wrong', '\tImpact and what we are doing about it', 'Issue: what is already wrong', '\tWhat it blocks and who is fixing it']);
      add('text', 'Decisions Needed', ['What you need from the audience', 'By when, and what happens if it slips']);
    },
    meeting(add, pres) {
      add('title', 'Team Meeting', ['Date · Time · Room or video link']);
      diagram(pres, add('text', 'Agenda'), 'vProcess', ['Welcome and check-in', 'Updates from last week', 'Topic for discussion', 'Action items and next steps']);
      add('twoText', 'Updates', ['Wins', '\tWhat went well', '\tWho to thank'], ['Blockers', '\tWhat is in the way', '\tWhat help is needed']);
      add('text', 'Discussion', ['The question to decide today', '\tOption A', '\tOption B']);
      add('text', 'Action Items', ['Action — owner — due date', 'Action — owner — due date', 'Action — owner — due date']);
      add('title', 'Thank You', ['Next meeting: date and time']);
    },
    lesson(add, pres) {
      add('title', 'Lesson Title', ['Course name · Teacher name']);
      add('text', 'Today We Will', ['Learn what …', 'Practise how to …', 'Be able to explain …']);
      diagram(pres, add('text', 'Key Idea'), 'radial', ['The idea', '\tExample', '\tWhy it matters', '\tCounter-example', '\tWhere it is used']);
      add('twoText', 'Compare', ['First thing', '\tFeature', '\tFeature'], ['Second thing', '\tFeature', '\tFeature']);
      add('text', 'Your Turn', ['A question or exercise for the class', 'How long they have, and in pairs or alone']);
      add('text', 'Summary and Homework', ['Three things to remember', 'Homework: what to do, due when']);
    },
    event(add, pres) {
      add('title', 'Event Name', ['Day, date · Time', 'Place']);
      add('text', 'What to Expect', ['The highlight of the event', 'Who is speaking or performing', 'Food, drinks and activities']);
      diagram(pres, add('text', 'Programme'), 'timeline', ['Time — opening', 'Time — main part', 'Time — break', 'Time — closing']);
      add('text', 'Getting There', ['Address', 'Public transport and parking', 'Accessibility']);
      add('title', 'See You There', ['RSVP by date · contact']);
    },
  };

  /** a new presentation from the template; size: { w, h } (default 4:3) */
  TL.build = function (id, size) {
    const t = LIST.find((x) => x.id === id);
    if (!t) return null;
    const pres = M.newPresentation(Object.assign({ design: t.design, empty: true }, size || {}));
    const did = Object.keys(pres.designs)[0];
    /** add(layout, title, body lines[, second column lines]): '\t' indents a line one level */
    const add = (layout, title, ...cols) => {
      const s = M.newSlide(pres, layout, did);
      pres.slides.push(s);
      const ph = (types) => s.shapes.filter((x) => x.ph && types.includes(x.ph.type));
      const tt = ph(['ctrTitle', 'title'])[0];
      if (tt) T.setPlain(tt.tx, title);
      if (layout === 'title') { const sub = ph(['subTitle'])[0]; if (sub && cols[0]) sub.tx.ps = cols[0].map((ln) => T.para(ln)); }
      else ph(['body']).forEach((b, i) => { if (cols[i]) b.tx.ps = T.fromLines(cols[i]); });
      s.trans = { type: 'fade', spd: 'med', click: true, after: null };
      return s;
    };
    B[id](add, pres);
    pres.props.title = '';
    pres.hf = Object.assign(pres.hf, { num: true, notOnTitle: true });
    return pres;
  };
})();
