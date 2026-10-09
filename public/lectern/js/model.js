/* Lectern — presentation model.
 * Units are points. Colours are '#RRGGBB' literals or scheme slot names
 * (dk1, lt1, dk2, lt2, tx1, bg1, tx2, bg2, accent1..6, hlink, folHlink).
 */
(function () {
  'use strict';
  const L = window.L;
  const M = (L.model = {});

  /* ---------- default text levels (PowerPoint 2003 master) ---------- */
  const BODY_BULLETS = ['•', '–', '•', '–', '»', '•', '•', '•', '•'];
  const BODY_SIZES = [32, 28, 24, 20, 20, 20, 20, 20, 20];
  const BODY_MARL = [27, 58.5, 90, 126, 162, 198, 234, 270, 306];
  const BODY_INDENT = [-27, -22.5, -18, -18, -18, -18, -18, -18, -18];
  M.DEFAULT_TX = {
    title: [{ algn: 'ctr', marL: 0, indent: 0, lnSpc: { pct: 100 }, spcBef: { pct: 0 }, spcAft: { pct: 0 }, bu: { t: 'none' }, rPr: { font: '+mj', sz: 44, color: 'tx2' } }],
    body: BODY_SIZES.map((sz, i) => ({
      algn: 'l', marL: BODY_MARL[i], indent: BODY_INDENT[i], lnSpc: { pct: 100 }, spcBef: { pct: 20 }, spcAft: { pct: 0 },
      bu: { t: 'char', ch: BODY_BULLETS[i] }, rPr: { font: '+mn', sz, color: 'tx1' },
    })),
    other: Array.from({ length: 9 }, (_, i) => ({
      algn: 'l', marL: i * 36, indent: 0, lnSpc: { pct: 100 }, spcBef: { pct: 0 }, spcAft: { pct: 0 }, bu: { t: 'none' }, rPr: { font: '+mn', sz: 18, color: 'tx1' },
    })),
  };

  /* ---------- colour schemes ----------
   * 2003 slot names in brackets: bg=lt1 [Background], tx1=dk1 [Text and lines], lt2 [Shadows], dk2 [Title text],
   * accent1 [Fills], accent2 [Accent], hlink [Accent and hyperlink], folHlink [Accent and followed hyperlink].
   */
  const scheme = (lt1, dk1, lt2, dk2, a1, a2, hl, fh, a3, a4, a5, a6) => ({
    lt1, dk1, lt2, dk2, accent1: a1, accent2: a2, hlink: hl, folHlink: fh,
    accent3: a3 || L.color.mix(a1, lt1, 0.35), accent4: a4 || L.color.mix(dk1, lt1, 0.3), accent5: a5 || L.color.mix(a2, lt1, 0.3), accent6: a6 || L.color.mix(hl, dk1, 0.2),
  });
  M.COLOR_SCHEMES = [
    scheme('#FFFFFF', '#000000', '#808080', '#000000', '#BBE0E3', '#333399', '#009999', '#99CC00', '#FFFFFF', '#000000', '#ADADD6', '#008A8A'),
    scheme('#FFFFFF', '#000000', '#969696', '#003366', '#3399FF', '#99CCFF', '#CC0000', '#FFCC00'),
    scheme('#000066', '#FFFFFF', '#000033', '#FFFF99', '#3366CC', '#00CCFF', '#9999FF', '#FF9900'),
    scheme('#FFFFEB', '#000000', '#808080', '#663300', '#CC9900', '#CC3300', '#669900', '#999966'),
    scheme('#003300', '#FFFFFF', '#001A00', '#CCFFCC', '#669966', '#99CC66', '#FFCC66', '#FF9933'),
    scheme('#333333', '#FFFFFF', '#1A1A1A', '#FFCC66', '#5C8AB8', '#E37B33', '#8DBF57', '#B5B5B5'),
    scheme('#F0F0F0', '#333333', '#999999', '#7A1F2B', '#C0504D', '#4F81BD', '#9BBB59', '#8064A2'),
    scheme('#FDFBF4', '#3B2A1A', '#A69276', '#5A3A1A', '#C8A96B', '#8C5A2B', '#2E6E8E', '#7A6A55'),
    scheme('#EAF3FA', '#0B2545', '#8DA9C4', '#13315C', '#5FA8D3', '#E07A5F', '#3D5A80', '#98C1D9'),
    scheme('#1B1F2A', '#E6E9F0', '#0E1118', '#7FDBCA', '#4CC9F0', '#F72585', '#B5179E', '#4361EE'),
  ];

  /* ---------- helpers for decoration shapes ---------- */
  let decoId = 0;
  const deco = (geom, x, y, w, h, fill, extra) => Object.assign({ id: 'd' + (decoId++), type: 'shape', name: 'Decoration', geom, x, y, w, h, rot: 0, fill, line: { t: 'none' } }, extra || {});
  const solid = (c, a) => ({ t: 'solid', c, a: a == null ? 1 : a });
  const grad = (stops, ang, path) => ({ t: 'grad', stops: stops.map(([p, c, a]) => ({ p, c, a: a == null ? 1 : a })), ang: ang || 0, path: path || 'lin' });
  const custom = (x, y, w, h, cmds, fill, extra) => deco('custom', x, y, w, h, fill, Object.assign({ path: { paths: [{ w, h, cmds, fill: 'norm', stroke: false }] } }, extra || {}));
  M.deco = deco; M.solid = solid; M.grad = grad;

  /* placeholder frames as fractions of slide width/height (2003 default master) */
  const PH = {
    title: [0.05, 0.04, 0.9, 0.1667], body: [0.05, 0.2333, 0.9, 0.66],
    ctrTitle: [0.075, 0.3106, 0.85, 0.2144], subTitle: [0.15, 0.5667, 0.7, 0.2556],
    dt: [0.05, 0.9104, 0.2333, 0.0694], ftr: [0.3417, 0.9104, 0.3167, 0.0694], sldNum: [0.7167, 0.9104, 0.2333, 0.0694],
  };

  /* ---------- design templates (original artwork) ---------- */
  const DESIGNS = [];
  const defDesign = (key, name, o) => DESIGNS.push(Object.assign({ key, name }, o));

  defDesign('default', 'Default Design', { colors: M.COLOR_SCHEMES[0], fonts: { major: 'Arial', minor: 'Arial' }, bg: solid('bg1') });

  defDesign('azure', 'Azure Horizon', {
    colors: scheme('#FFFFFF', '#1B2A41', '#7F93B2', '#FFFFFF', '#4A7BC8', '#E0A526', '#2D6BB5', '#8E7CC3', '#9BB7E3', '#1F3A6E', '#F2C76E', '#5C88C5'),
    fonts: { major: 'Trebuchet MS', minor: 'Verdana' },
    bg: grad([[0, '#FFFFFF'], [1, '#DCE8F7']], 90),
    deco: (W, H) => [deco('rect', 0, 0, W, H * 0.2, grad([[0, '#2A4E8F'], [1, '#1F3A6E']], 90)), deco('rect', 0, H * 0.2, W, H * 0.008, solid('accent2'))],
    titleDeco: (W, H) => [deco('rect', 0, H * 0.3, W, H * 0.26, grad([[0, '#2A4E8F'], [1, '#1F3A6E']], 90)), deco('rect', 0, H * 0.56, W, H * 0.008, solid('accent2'))],
    ph: { title: [0.05, 0.025, 0.9, 0.15], body: [0.07, 0.26, 0.86, 0.63], ctrTitle: [0.075, 0.31, 0.85, 0.24], subTitle: [0.15, 0.6, 0.7, 0.2] },
    tx: { title: { rPr: { color: 'lt1', b: true } }, body: { bu: { c: 'accent2' } } },
  });

  defDesign('night', 'Night Sky', {
    colors: scheme('#0B1633', '#FFFFFF', '#06102A', '#FFE599', '#3D6CB3', '#FFD966', '#9FC5F8', '#C9DAF8', '#6D9EEB', '#B4A7D6', '#F6B26B', '#93C47D'),
    fonts: { major: 'Georgia', minor: 'Verdana' },
    bg: grad([[0, '#0B1633'], [1, '#22427A']], 90),
    deco: (W, H) => {
      const stars = [[0.82, 0.06, 0.022], [0.9, 0.14, 0.014], [0.76, 0.16, 0.01], [0.94, 0.05, 0.012], [0.68, 0.05, 0.008], [0.87, 0.24, 0.009], [0.06, 0.92, 0.01], [0.13, 0.86, 0.007]];
      return [
        ...stars.map(([x, y, s]) => deco('star4', W * x, H * y, W * s * 1.6, W * s * 1.6, solid('#FFF2CC', 0.9))),
        deco('ellipse', W * 0.62, H * 0.88, W * 0.7, H * 0.5, solid('#FFFFFF', 0.04)),
      ];
    },
    tx: { title: { algn: 'l' }, body: { bu: { c: 'accent2', ch: '✦' } } },
  });

  defDesign('meadow', 'Meadow', {
    colors: scheme('#F7FBEF', '#22331A', '#A7C08F', '#2F5E1E', '#7DB04B', '#E7A93B', '#3B7D2A', '#8A9A5B', '#B9D88F', '#4E7A33', '#F3D27A', '#6A9C89'),
    fonts: { major: 'Century Gothic', minor: 'Century Gothic' },
    bg: solid('bg1'),
    deco: (W, H) => [
      custom(0, H * 0.78, W, H * 0.22, [['M', 0, H * 0.12], ['C', W * 0.25, -H * 0.02, W * 0.45, H * 0.18, W * 0.7, H * 0.06], ['C', W * 0.85, 0, W * 0.95, H * 0.04, W, H * 0.07], ['L', W, H * 0.22], ['L', 0, H * 0.22], ['Z']], solid('#9CCB5E', 0.55)),
      custom(0, H * 0.83, W, H * 0.17, [['M', 0, H * 0.1], ['C', W * 0.3, H * 0.02, W * 0.55, H * 0.14, W, 0], ['L', W, H * 0.17], ['L', 0, H * 0.17], ['Z']], solid('#5E9E3A', 0.65)),
    ],
    ph: { body: [0.05, 0.2333, 0.9, 0.56], dt: [0.05, 0.92, 0.2333, 0.06], ftr: [0.3417, 0.92, 0.3167, 0.06], sldNum: [0.7167, 0.92, 0.2333, 0.06] },
    tx: { title: { algn: 'l' }, body: { bu: { c: 'accent1', ch: '●' } }, footer: { color: '#FFFFFF' } },
  });

  defDesign('ember', 'Ember', {
    colors: scheme('#FFF8F0', '#3A2414', '#C9A88C', '#8C2D04', '#E8641C', '#C1272D', '#B5651D', '#8E6C4F', '#F4A259', '#5B3A29', '#F7C59F', '#A23B05'),
    fonts: { major: 'Franklin Gothic Medium', minor: 'Arial' },
    bg: solid('bg1'),
    deco: (W, H) => [
      deco('rect', 0, H * 0.94, W, H * 0.06, grad([[0, '#C1272D'], [1, '#F39C12']], 0)),
      deco('rect', 0, H * 0.925, W, H * 0.008, solid('accent1', 0.7)),
      deco('rect', 0, H * 0.905, W, H * 0.005, solid('accent1', 0.4)),
      deco('rect', W * 0.05, H * 0.205, W * 0.9, H * 0.004, solid('accent1')),
    ],
    titleDeco: (W, H) => [
      deco('rect', 0, H * 0.94, W, H * 0.06, grad([[0, '#C1272D'], [1, '#F39C12']], 0)),
      deco('rect', 0, 0, W * 0.035, H * 0.94, grad([[0, '#F39C12'], [1, '#C1272D']], 90)),
      deco('rect', W * 0.1, H * 0.54, W * 0.8, H * 0.005, solid('accent1')),
    ],
    ph: { body: [0.05, 0.24, 0.9, 0.64], dt: [0.05, 0.86, 0.2333, 0.045], ftr: [0.3417, 0.86, 0.3167, 0.045], sldNum: [0.7167, 0.86, 0.2333, 0.045] },
    tx: { title: { algn: 'l' }, body: { bu: { c: 'accent1', ch: '■' } } },
  });

  defDesign('graphite', 'Graphite', {
    colors: scheme('#2B2D31', '#E8EAED', '#17181B', '#FFFFFF', '#F28C28', '#58A6FF', '#79C0FF', '#B1BAC4', '#FFB86B', '#8B949E', '#3FB950', '#D2A8FF'),
    fonts: { major: 'Arial', minor: 'Arial' },
    bg: grad([[0, '#34373D'], [1, '#232428']], 90),
    deco: (W, H) => [
      deco('parallelogram', -W * 0.04, H * 0.03, W * 0.24, H * 0.022, solid('accent1'), { adj: { adj: 60000 } }),
      deco('rect', W * 0.05, H * 0.215, W * 0.9, H * 0.0025, solid('#FFFFFF', 0.25)),
    ],
    tx: { title: { algn: 'l', rPr: { b: true } }, body: { bu: { c: 'accent1', ch: '▸' } } },
  });

  defDesign('blueprint', 'Blueprint', {
    colors: scheme('#1D4E89', '#E8F1FB', '#123359', '#FFFFFF', '#7FB3E6', '#FFD24D', '#B3D4F5', '#9DBCE0', '#4A86C5', '#C9DDF2', '#FFE699', '#5FA0DA'),
    fonts: { major: 'Courier New', minor: 'Arial' },
    bg: solid('bg1'),
    deco: (W, H) => [
      deco('rect', 0, 0, W, H, { t: 'patt', prst: 'smGrid', fg: '#3D6FA8', bg: '#1D4E89' }),
      deco('rect', W * 0.03, H * 0.04, W * 0.94, H * 0.92, { t: 'none' }, { line: { c: '#B3D4F5', w: 1, dash: 'solid' } }),
    ],
    tx: { title: { algn: 'l', rPr: { b: true } }, body: { bu: { c: 'accent2', ch: '–' } } },
  });

  defDesign('capsule', 'Capsule', {
    colors: scheme('#FFFFFF', '#1E2D33', '#9FB4BA', '#21545F', '#2BA6A0', '#9BC53D', '#1F7A8C', '#7E9C9F', '#7FD1C8', '#33658A', '#E3D96E', '#F26B38'),
    fonts: { major: 'Trebuchet MS', minor: 'Trebuchet MS' },
    bg: solid('bg1'),
    deco: (W, H) => [
      deco('roundRect', W * 0.68, H * 0.035, W * 0.2, H * 0.05, solid('accent1'), { adj: { adj: 50000 } }),
      deco('roundRect', W * 0.84, H * 0.035, W * 0.12, H * 0.05, solid('accent2'), { adj: { adj: 50000 } }),
      deco('roundRect', W * 0.76, H * 0.1, W * 0.18, H * 0.05, solid('accent1', 0.45), { adj: { adj: 50000 } }),
      deco('rect', 0, H * 0.965, W, H * 0.035, solid('accent1')),
    ],
    ph: { title: [0.05, 0.06, 0.62, 0.15] },
    tx: { title: { algn: 'l' }, body: { bu: { c: 'accent1', ch: '●' } } },
  });

  defDesign('ocean', 'Ocean Wave', {
    colors: scheme('#EEF8FB', '#0E2F3A', '#8FB9C4', '#0E4C5C', '#2A9D8F', '#F4A261', '#1D7C99', '#7A9EA8', '#8ECAE6', '#264653', '#E9C46A', '#219EBC'),
    fonts: { major: 'Tahoma', minor: 'Tahoma' },
    bg: grad([[0, '#F4FBFD'], [1, '#C7E6EE']], 90),
    deco: (W, H) => [
      custom(0, H * 0.84, W, H * 0.16, [['M', 0, H * 0.06], ['C', W * 0.2, -H * 0.03, W * 0.35, H * 0.12, W * 0.55, H * 0.05], ['C', W * 0.75, -H * 0.02, W * 0.88, H * 0.08, W, H * 0.03], ['L', W, H * 0.16], ['L', 0, H * 0.16], ['Z']], solid('#2A9D8F', 0.45)),
      custom(0, H * 0.89, W, H * 0.11, [['M', 0, H * 0.05], ['C', W * 0.25, H * 0.1, W * 0.45, 0, W * 0.7, H * 0.04], ['C', W * 0.85, H * 0.07, W * 0.95, H * 0.02, W, H * 0.03], ['L', W, H * 0.11], ['L', 0, H * 0.11], ['Z']], solid('#1D7C99', 0.6)),
    ],
    ph: { body: [0.05, 0.2333, 0.9, 0.58], dt: [0.05, 0.925, 0.2333, 0.06], ftr: [0.3417, 0.925, 0.3167, 0.06], sldNum: [0.7167, 0.925, 0.2333, 0.06] },
    tx: { body: { bu: { c: 'accent1', ch: '•' } }, footer: { color: '#FFFFFF' } },
  });

  defDesign('parchment', 'Parchment', {
    colors: scheme('#F4ECD8', '#3B2A1A', '#B8A27F', '#5A3A1A', '#C8A96B', '#8C5A2B', '#7B3F00', '#8A7A60', '#E1CDA0', '#6B4F2E', '#A77B4F', '#556B2F'),
    fonts: { major: 'Book Antiqua', minor: 'Georgia' },
    bg: grad([[0, '#F7F0DE'], [1, '#E6D7B5']], 0, 'circle'),
    deco: (W, H) => [
      deco('rect', W * 0.05, H * 0.205, W * 0.9, H * 0.0035, solid('accent2')),
      deco('rect', W * 0.05, H * 0.214, W * 0.9, H * 0.0015, solid('accent2')),
      deco('rect', W * 0.02, H * 0.027, W * 0.96, H * 0.946, { t: 'none' }, { line: { c: '#8C5A2B', w: 0.75, dash: 'solid' } }),
    ],
    tx: { body: { bu: { c: 'accent2', ch: '❧' } } },
  });

  defDesign('pixel', 'Pixel Mosaic', {
    colors: scheme('#FFFFFF', '#1A2733', '#A0AEBB', '#1F4E79', '#2E86C1', '#28B463', '#1B4F72', '#7F8C8D', '#85C1E9', '#17A589', '#F4D03F', '#AF7AC5'),
    fonts: { major: 'Verdana', minor: 'Verdana' },
    bg: solid('bg1'),
    deco: (W, H) => {
      const cells = [[0, 0, 'accent1'], [1, 0, 'accent3'], [2, 0, 'accent2'], [3, 0, 'accent1'], [1, 1, 'accent1'], [2, 1, 'accent6'], [3, 1, 'accent3'], [2, 2, 'accent2'], [3, 2, 'accent1'], [3, 3, 'accent3']];
      const s = W * 0.028, g = s * 0.18, x0 = W - (s + g) * 4 - W * 0.02, y0 = H * 0.03;
      return [...cells.map(([cx, cy, c]) => deco('rect', x0 + cx * (s + g), y0 + cy * (s + g), s, s, solid(c, 0.9))), deco('rect', W * 0.05, H * 0.21, W * 0.62, H * 0.004, solid('accent1'))];
    },
    ph: { title: [0.05, 0.04, 0.78, 0.1667] },
    tx: { title: { algn: 'l' }, body: { bu: { c: 'accent2', ch: '■' } } },
  });

  defDesign('crimson', 'Crimson Edge', {
    colors: scheme('#FFFFFF', '#262626', '#A6A6A6', '#A4161A', '#A4161A', '#495057', '#0B6E99', '#6C757D', '#E5383B', '#343A40', '#F4ACB7', '#ADB5BD'),
    fonts: { major: 'Georgia', minor: 'Arial' },
    bg: solid('bg1'),
    deco: (W, H) => [deco('rect', 0, 0, W * 0.03, H, solid('accent1')), deco('rect', W * 0.08, H * 0.205, W * 0.87, H * 0.0025, solid('accent1'))],
    titleDeco: (W, H) => [deco('rect', 0, 0, W * 0.03, H, solid('accent1')), deco('rect', W * 0.03, H * 0.6, W * 0.97, H * 0.006, solid('accent1'))],
    ph: { title: [0.08, 0.04, 0.87, 0.1667], body: [0.08, 0.2333, 0.87, 0.66], ctrTitle: [0.08, 0.34, 0.84, 0.24], subTitle: [0.08, 0.63, 0.84, 0.2], dt: [0.08, 0.9104, 0.2, 0.0694] },
    tx: { title: { algn: 'l' }, body: { bu: { c: 'accent1', ch: '▪' } }, ctrTitleAlgn: 'l', subTitleAlgn: 'l' },
  });

  defDesign('glass', 'Layered Glass', {
    colors: scheme('#3A6EA5', '#FFFFFF', '#244A73', '#FFFFFF', '#9CC3EB', '#FFD166', '#CFE4FA', '#B8D0EA', '#5D93CC', '#E3F0FF', '#FFE29A', '#7AB3E6'),
    fonts: { major: 'Tahoma', minor: 'Tahoma' },
    bg: grad([[0, '#2F5F94'], [1, '#6A9FD4']], 45),
    deco: (W, H) => [
      deco('rect', W * 0.55, -H * 0.1, W * 0.6, H * 0.7, solid('#FFFFFF', 0.08), { rot: 18 }),
      deco('rect', W * 0.7, H * 0.4, W * 0.5, H * 0.8, solid('#FFFFFF', 0.06), { rot: -12 }),
      deco('rect', -W * 0.15, H * 0.75, W * 0.6, H * 0.4, solid('#FFFFFF', 0.07), { rot: 8 }),
    ],
    tx: { body: { bu: { c: 'accent2' } } },
  });

  defDesign('radial', 'Radial Glow', {
    colors: scheme('#0D1F3C', '#DCE6F2', '#081428', '#FFFFFF', '#7FB3FF', '#FFC857', '#A9CBFF', '#8FA3BF', '#4F7CC2', '#C3D5EE', '#FF8C42', '#6FD6C6'),
    fonts: { major: 'Arial', minor: 'Arial' },
    bg: grad([[0, '#2E5C9A'], [1, '#0D1F3C']], 0, 'circle'),
    deco: (W, H) => [deco('ellipse', W * 0.3, H * 0.95, W * 0.4, H * 0.012, solid('#7FB3FF', 0.35))],
    tx: { body: { bu: { c: 'accent1' } } },
  });

  M.DESIGNS = DESIGNS;

  /** Build a concrete design object (data only) for a slide size. */
  M.buildDesign = function (key, W, H) {
    const t = DESIGNS.find((d) => d.key === key) || DESIGNS[0];
    decoId = 0;
    const phFrac = Object.assign({}, PH, t.ph || {});
    const ph = {};
    for (const k in phFrac) { const [x, y, w, h] = phFrac[k]; ph[k] = { x: L.round(W * x, 2), y: L.round(H * y, 2), w: L.round(W * w, 2), h: L.round(H * h, 2) }; }
    const txs = L.clone(M.DEFAULT_TX);
    const o = t.tx || {};
    if (o.title) txs.title[0] = L.deepMerge(txs.title[0], o.title);
    if (o.body) txs.body = txs.body.map((lv, i) => L.deepMerge(lv, (i === 1 || i === 3) && o.body.bu && o.body.bu.ch ? Object.assign({}, o.body, { bu: Object.assign({}, o.body.bu, { ch: '–' }) }) : o.body));
    const footer = { algn: 'l', rPr: { sz: 14, color: (o.footer && o.footer.color) || 'tx1' } };
    /* text sizes are designed for a 10 x 7.5 in slide; smaller slides get proportionally smaller text */
    const tk = Math.min(1, W / 720, H / 540);
    if (tk < 0.999) {
      const sc = (st) => { if (st.rPr && st.rPr.sz) st.rPr.sz = L.round(st.rPr.sz * tk, 1); if (st.marL) st.marL = L.round(st.marL * tk, 2); if (st.indent) st.indent = L.round(st.indent * tk, 2); };
      for (const cls in txs) txs[cls].forEach(sc);
      footer.rPr.sz = L.round(footer.rPr.sz * tk, 1);
    }
    const d = {
      id: L.uid('dsn'), key: t.key, name: t.name,
      colors: L.clone(t.colors), fonts: L.clone(t.fonts),
      bg: L.clone(t.bg), deco: t.deco ? t.deco(W, H) : [],
      titleDeco: t.titleDeco ? t.titleDeco(W, H) : null, titleBg: t.titleBg ? L.clone(t.titleBg) : null,
      ph, tx: txs, footer,
      ctrTitleAlgn: o.ctrTitleAlgn || 'ctr', subTitleAlgn: o.subTitleAlgn || 'ctr',
    };
    return d;
  };

  /* ---------- colour resolution ---------- */
  const MAP = { tx1: 'dk1', bg1: 'lt1', tx2: 'dk2', bg2: 'lt2' };
  M.resolveColor = function (c, design, fallback) {
    if (c == null || c === '') return fallback == null ? '#000000' : fallback;
    if (c[0] === '#') return c;
    const cols = (design && design.colors) || M.COLOR_SCHEMES[0];
    const slot = (design && design.clrMap && design.clrMap[c]) || MAP[c] || c;
    return cols[slot] || L.color.PRESET[c] || fallback || '#000000';
  };
  M.SCHEME_SLOTS = [['lt1', 'Background'], ['dk1', 'Text and lines'], ['lt2', 'Shadows'], ['dk2', 'Title text'], ['accent1', 'Fills'], ['accent2', 'Accent'], ['hlink', 'Accent and hyperlink'], ['folHlink', 'Accent and followed hyperlink']];
  M.schemeRow = (design) => ['bg1', 'tx1', 'bg2', 'tx2', 'accent1', 'accent2', 'hlink', 'folHlink'].map((k) => ({ ref: k, hex: M.resolveColor(k, design) }));

  /* ---------- text model ---------- */
  const T = (L.txt = {});
  T.body = (paras, o) => Object.assign({ anchor: 't', wrap: true, autofit: 'none', ins: [7.2, 3.6, 7.2, 3.6], ps: paras || [T.para('')] }, o || {});
  T.para = (text, pp, rp, lvl) => {
    const runs = [];
    const lines = String(text == null ? '' : text);
    if (lines.length) runs.push(Object.assign({ t: lines }, rp || {}));
    return { lvl: lvl || 0, pp: pp ? L.clone(pp) : {}, rs: runs, end: rp ? L.clone(rp) : undefined };
  };
  T.fromLines = (lines, pp, rp) => lines.map((ln) => {
    if (typeof ln === 'object' && ln) return T.para(ln.t, Object.assign({}, pp || {}, ln.pp || {}), Object.assign({}, rp || {}, ln.rp || {}), ln.lvl || 0);
    const m = /^(\t*)(.*)$/.exec(ln);
    return T.para(m[2], pp, rp, m[1].length);
  });
  T.paraText = (p) => p.rs.map((r) => (r.fld ? r.t || '' : r.t)).join('');
  T.plain = (tx) => (tx && tx.ps ? tx.ps.map(T.paraText).join('\n') : '');
  T.isEmpty = (tx) => !tx || !tx.ps || tx.ps.every((p) => !T.paraText(p).length);
  T.RUN_KEYS = ['b', 'i', 'u', 'strike', 'sz', 'font', 'color', 'base', 'shd', 'shdX', 'link', 'fld', 'spc', 'cap', 'fill', 'ln', 'hl', 'emb', 'keep'];
  T.runProps = (r) => { const o = {}; for (const k of T.RUN_KEYS) if (r[k] !== undefined && k !== 'fld') o[k] = r[k]; return o; };
  T.sameProps = (a, b) => L.equal(T.runProps(a), T.runProps(b)) && !a.fld && !b.fld;
  T.normalize = (p) => {
    const out = [];
    for (const r of p.rs) {
      if (!r.fld && !r.t && !r.keep) continue;
      const prev = out[out.length - 1];
      if (prev && T.sameProps(prev, r)) prev.t += r.t;
      else out.push(Object.assign({}, r));
    }
    if (!out.length && p.rs.length) p.end = Object.assign({}, p.end || {}, T.runProps(p.rs[0]));
    p.rs = out;
    return p;
  };
  /* split runs of a paragraph at character offset; returns index of first run at/after offset */
  T.splitAt = (p, off) => {
    let pos = 0;
    for (let i = 0; i < p.rs.length; i++) {
      const r = p.rs[i];
      const len = r.t.length;
      if (off === pos) return i;
      if (off < pos + len) {
        if (r.fld) return i; /* fields are atomic */
        const a = Object.assign({}, r, { t: r.t.slice(0, off - pos) });
        const b = Object.assign({}, r, { t: r.t.slice(off - pos) });
        p.rs.splice(i, 1, a, b);
        return i + 1;
      }
      pos += len;
    }
    return p.rs.length;
  };
  /* apply fn to runs within [start,end) of paragraph selection {p0,o0,p1,o1} */
  T.mapRange = (tx, sel, fn) => {
    for (let pi = sel.p0; pi <= sel.p1; pi++) {
      const p = tx.ps[pi];
      if (!p) continue;
      const len = T.paraText(p).length;
      const a = pi === sel.p0 ? sel.o0 : 0, b = pi === sel.p1 ? sel.o1 : len;
      if (len === 0) { p.end = fn(Object.assign({}, p.end || {})); continue; }
      if (a === b && sel.p0 !== sel.p1) continue;
      const ia = T.splitAt(p, a), ib = T.splitAt(p, b);
      for (let i = ia; i < ib; i++) p.rs[i] = fn(p.rs[i]);
      T.normalize(p);
    }
  };
  T.allRange = (tx) => ({ p0: 0, o0: 0, p1: tx.ps.length - 1, o1: T.paraText(tx.ps[tx.ps.length - 1]).length });
  T.setPlain = (tx, str) => {
    const first = tx.ps[0] || T.para('');
    const rp = first.rs[0] ? T.runProps(first.rs[0]) : first.end || {};
    tx.ps = String(str).split('\n').map((line, i) => {
      const src = tx.ps[i] || first;
      return { lvl: src.lvl || 0, pp: L.clone(src.pp || {}), rs: line ? [Object.assign({ t: line }, rp)] : [], end: L.clone(rp) };
    });
  };

  /* ---------- style resolution ---------- */
  const S = (L.style = {});
  S.cls = (sh) => {
    if (!sh || !sh.ph || sh.type === 'table' || sh.type === 'chart' || sh.type === 'image') return 'other';
    const t = sh.ph.type;
    if (t === 'title' || t === 'ctrTitle') return 'title';
    if (t === 'dt' || t === 'ftr' || t === 'sldNum' || t === 'hdr') return 'other';
    return 'body';
  };
  S.level = (sh, lvl, design) => {
    const cls = S.cls(sh);
    lvl = L.clamp(lvl || 0, 0, 8);
    const base = M.DEFAULT_TX[cls][Math.min(lvl, M.DEFAULT_TX[cls].length - 1)];
    const ds = design && design.tx && design.tx[cls] ? design.tx[cls][Math.min(lvl, design.tx[cls].length - 1)] : null;
    let st = L.deepMerge(base, ds);
    if (sh && sh.ph && (sh.ph.type === 'dt' || sh.ph.type === 'ftr' || sh.ph.type === 'sldNum') && design && design.footer) {
      st = L.deepMerge(st, design.footer);
      st.algn = sh.ph.type === 'dt' ? 'l' : sh.ph.type === 'ftr' ? 'ctr' : 'r';
    }
    if (sh && sh.tx && sh.tx.lst && sh.tx.lst[lvl]) st = L.deepMerge(st, sh.tx.lst[lvl]);
    return st;
  };
  S.para = (sh, p, design) => {
    const st = S.level(sh, p.lvl, design);
    return p.pp ? L.deepMerge(st, p.pp) : st;
  };
  S.font = (f, design) => {
    if (!f) return (design && design.fonts && design.fonts.minor) || 'Arial';
    if (f === '+mj' || f === '+mj-lt') return (design && design.fonts && design.fonts.major) || 'Arial';
    if (f === '+mn' || f === '+mn-lt') return (design && design.fonts && design.fonts.minor) || 'Arial';
    return f;
  };
  S.run = (ps, r) => Object.assign({}, ps.rPr || {}, T.runProps(r || {}));

  /* ---------- layouts ---------- */
  M.LAYOUTS = [
    { key: 'title', name: 'Title Slide', group: 'Text Layouts', ox: 'title' },
    { key: 'titleOnly', name: 'Title Only', group: 'Text Layouts', ox: 'titleOnly' },
    { key: 'text', name: 'Title and Text', group: 'Text Layouts', ox: 'tx' },
    { key: 'twoText', name: 'Title and 2-Column Text', group: 'Text Layouts', ox: 'twoColTx' },
    { key: 'blank', name: 'Blank', group: 'Content Layouts', ox: 'blank' },
    { key: 'contentOnly', name: 'Content', group: 'Content Layouts', ox: 'objOnly' },
    { key: 'content', name: 'Title and Content', group: 'Content Layouts', ox: 'obj' },
    { key: 'twoContent', name: 'Title and 2 Content', group: 'Content Layouts', ox: 'twoObj' },
    { key: 'contentTwoContent', name: 'Title, Content and 2 Content', group: 'Content Layouts', ox: 'objAndTwoObj' },
    { key: 'twoContentContent', name: 'Title, 2 Content and Content', group: 'Content Layouts', ox: 'twoObjAndObj' },
    { key: 'fourContent', name: 'Title and 4 Content', group: 'Content Layouts', ox: 'fourObj' },
    { key: 'textContent', name: 'Title, Text, and Content', group: 'Text and Content Layouts', ox: 'txAndObj' },
    { key: 'contentText', name: 'Title, Content and Text', group: 'Text and Content Layouts', ox: 'objAndTx' },
    { key: 'textTwoContent', name: 'Title, Text, and 2 Content', group: 'Text and Content Layouts', ox: 'txAndTwoObj' },
    { key: 'twoContentText', name: 'Title, 2 Content and Text', group: 'Text and Content Layouts', ox: 'twoObjAndTx' },
    { key: 'table', name: 'Title and Table', group: 'Other Layouts', ox: 'tbl' },
    { key: 'chart', name: 'Title and Chart', group: 'Other Layouts', ox: 'chart' },
    { key: 'textOverContent', name: 'Title and Text over Content', group: 'Other Layouts', ox: 'txOverObj' },
    { key: 'contentOverText', name: 'Title and Content over Text', group: 'Other Layouts', ox: 'objOverTx' },
    { key: 'vertText', name: 'Title and Vertical Text', group: 'Other Layouts', ox: 'vertTx' },
    { key: 'vertTitleText', name: 'Vertical Title and Text', group: 'Other Layouts', ox: 'vertTitleAndTx' },
  ];
  M.layoutInfo = (key) => M.LAYOUTS.find((l) => l.key === key) || M.LAYOUTS[2];

  /** Placeholder frames for a layout: [{type, idx, x,y,w,h, vert?}] */
  M.layoutFrames = function (key, design) {
    const t = design.ph.title, b = design.ph.body;
    const gap = Math.round(b.w * 0.0185 * 100) / 100;
    const half = (b.w - gap) / 2, vh = (b.h - gap) / 2;
    const L1 = { x: b.x, y: b.y, w: half, h: b.h }, R1 = { x: b.x + half + gap, y: b.y, w: half, h: b.h };
    const LT = { x: b.x, y: b.y, w: half, h: vh }, LB = { x: b.x, y: b.y + vh + gap, w: half, h: vh };
    const RT = { x: b.x + half + gap, y: b.y, w: half, h: vh }, RB = { x: b.x + half + gap, y: b.y + vh + gap, w: half, h: vh };
    const TOP = { x: b.x, y: b.y, w: b.w, h: vh }, BOT = { x: b.x, y: b.y + vh + gap, w: b.w, h: vh };
    const ttl = Object.assign({ type: 'title' }, t);
    const f = (type, idx, r, extra) => Object.assign({ type, idx }, r, extra || {});
    switch (key) {
      case 'title': return [Object.assign({ type: 'ctrTitle' }, design.ph.ctrTitle), f('subTitle', 1, design.ph.subTitle)];
      case 'titleOnly': return [ttl];
      case 'text': return [ttl, f('body', 1, b)];
      case 'twoText': return [ttl, f('body', 1, L1), f('body', 2, R1)];
      case 'blank': return [];
      case 'contentOnly': return [f('obj', 1, { x: b.x, y: t.y, w: b.w, h: b.y + b.h - t.y })];
      case 'content': return [ttl, f('obj', 1, b)];
      case 'twoContent': return [ttl, f('obj', 1, L1), f('obj', 2, R1)];
      case 'contentTwoContent': return [ttl, f('obj', 1, L1), f('obj', 2, RT), f('obj', 3, RB)];
      case 'twoContentContent': return [ttl, f('obj', 1, LT), f('obj', 2, LB), f('obj', 3, R1)];
      case 'fourContent': return [ttl, f('obj', 1, LT), f('obj', 2, RT), f('obj', 3, LB), f('obj', 4, RB)];
      case 'textContent': return [ttl, f('body', 1, L1), f('obj', 2, R1)];
      case 'contentText': return [ttl, f('obj', 1, L1), f('body', 2, R1)];
      case 'textTwoContent': return [ttl, f('body', 1, L1), f('obj', 2, RT), f('obj', 3, RB)];
      case 'twoContentText': return [ttl, f('obj', 1, LT), f('obj', 2, LB), f('body', 3, R1)];
      case 'table': return [ttl, f('tbl', 1, b)];
      case 'chart': return [ttl, f('chart', 1, b)];
      case 'textOverContent': return [ttl, f('body', 1, TOP), f('obj', 2, BOT)];
      case 'contentOverText': return [ttl, f('obj', 1, TOP), f('body', 2, BOT)];
      case 'vertText': return [ttl, f('body', 1, b, { vert: 'vert' })];
      case 'vertTitleText': {
        const tw = b.w * 0.22;
        return [Object.assign({ type: 'title', vert: 'vert' }, { x: b.x + b.w - tw, y: t.y, w: tw, h: b.y + b.h - t.y }), f('body', 1, { x: b.x, y: t.y, w: b.w - tw - gap, h: b.y + b.h - t.y }, { vert: 'vert' })];
      }
      default: return [ttl, f('body', 1, b)];
    }
  };

  M.PROMPTS = { title: 'Click to add title', ctrTitle: 'Click to add title', subTitle: 'Click to add subtitle', body: 'Click to add text', obj: 'Click to add text', tbl: 'Double click to add table', chart: 'Double click to add chart', pic: 'Click icon to add picture', dt: '', ftr: '', sldNum: '' };

  M.makePlaceholder = function (fr, design, idHint) {
    const isTitle = fr.type === 'title' || fr.type === 'ctrTitle';
    const sh = {
      id: L.uid('s'), type: 'text', name: phName(fr.type, fr.idx), ph: { type: fr.type, idx: fr.idx || 0 },
      x: fr.x, y: fr.y, w: fr.w, h: fr.h, rot: 0, geom: 'rect', fill: { t: 'none' }, line: { t: 'none' },
      tx: T.body([T.para('')], { anchor: isTitle ? 'ctr' : 't', autofit: isTitle ? 'none' : 'norm' }),
    };
    if (fr.vert) sh.tx.vert = fr.vert;
    if (fr.type === 'ctrTitle') sh.tx.lst = { 0: { algn: design.ctrTitleAlgn || 'ctr' } };
    if (fr.type === 'subTitle') { sh.tx.lst = { 0: { marL: 0, indent: 0, algn: design.subTitleAlgn || 'ctr', bu: { t: 'none' }, rPr: { color: 'tx1' } } }; sh.tx.anchor = 't'; }
    if (fr.type === 'vertTitle') sh.tx.vert = 'vert';
    return sh;
  };
  function phName(type, idx) {
    return ({ title: 'Title', ctrTitle: 'Title', subTitle: 'Subtitle', body: 'Text Placeholder', obj: 'Content Placeholder', tbl: 'Table Placeholder', chart: 'Chart Placeholder', dt: 'Date Placeholder', ftr: 'Footer Placeholder', sldNum: 'Slide Number Placeholder' }[type] || 'Placeholder') + ' ' + ((idx || 0) + 1);
  }

  /* ---------- presentation ---------- */
  M.SLIDE_SIZES = [
    { key: 'screen', name: 'On-screen Show (4:3)', w: 720, h: 540 },
    { key: 'widescreen', name: 'Widescreen (16:9), 13.33 x 7.5 in', w: 960, h: 540 },
    { key: 'wide', name: 'On-screen Show (16:9), 10 x 5.63 in', w: 720, h: 405 },
    { key: 'wide10', name: 'On-screen Show (16:10)', w: 720, h: 450 },
    { key: 'letter', name: 'Letter Paper (8.5x11 in)', w: 720, h: 540 },
    { key: 'a4', name: 'A4 Paper (210x297 mm)', w: 780, h: 540 },
    { key: '35mm', name: '35mm Slides', w: 810, h: 540 },
    { key: 'overhead', name: 'Overhead', w: 720, h: 540 },
    { key: 'banner', name: 'Banner', w: 576, h: 72 },
    { key: 'custom', name: 'Custom', w: 720, h: 540 },
  ];

  M.newPresentation = function (opts) {
    opts = opts || {};
    const W = opts.w || 720, H = opts.h || 540;
    const design = M.buildDesign(opts.design || 'default', W, H);
    const pres = {
      W, H, firstNum: 1, title: opts.title || '',
      designs: { [design.id]: design }, slides: [],
      show: { loop: false, noAnim: false, useTimings: true, from: 0, to: 0, penColor: '#FF0000', kiosk: false },
      hf: { dt: false, dtAuto: true, dtFmt: 'datetime1', dtText: '', num: false, ftr: false, ftrText: '', notOnTitle: false },
      props: { author: '', subject: '', keywords: '', comments: '', company: '', category: '', created: new Date().toISOString() },
    };
    if (opts.empty !== true) pres.slides.push(M.newSlide(pres, 'title', design.id));
    return pres;
  };

  M.newSlide = function (pres, layout, designId) {
    const design = pres.designs[designId] || Object.values(pres.designs)[0];
    const slide = {
      id: L.uid('sl'), layout: layout || 'text', design: design.id, shapes: [], notes: '',
      trans: { type: 'none', spd: 'fast', click: true, after: null }, anims: [], hidden: false,
    };
    for (const fr of M.layoutFrames(slide.layout, design)) slide.shapes.push(M.makePlaceholder(fr, design));
    return slide;
  };

  M.insertSlides = function (pres, at, slides) {
    L.preserve?.joinSection(pres, pres.slides, at, slides);
    pres.slides.splice(at, 0, ...slides);
  };
  M.design = (pres, slide) => (slide && pres.designs[slide.design]) || Object.values(pres.designs)[0];

  /** Re-apply a layout: keep existing content, move it into the new frames, add/remove empty placeholders. */
  M.applyLayout = function (pres, slide, key) {
    const design = M.design(pres, slide);
    const frames = M.layoutFrames(key, design);
    const old = slide.shapes;
    const used = new Set();
    const take = (pred) => { const s = old.find((x) => !used.has(x) && pred(x)); if (s) used.add(s); return s; };
    const out = [];
    for (const fr of frames) {
      const isTitle = fr.type === 'title' || fr.type === 'ctrTitle';
      let s = isTitle ? take((x) => x.ph && (x.ph.type === 'title' || x.ph.type === 'ctrTitle'))
        : take((x) => x.ph && x.ph.type !== 'title' && x.ph.type !== 'ctrTitle' && x.ph.type !== 'dt' && x.ph.type !== 'ftr' && x.ph.type !== 'sldNum');
      if (s && s.type === 'text') {
        const fresh = M.makePlaceholder(fr, design);
        const keep = s.tx;
        s.ph = fresh.ph; s.x = fr.x; s.y = fr.y; s.w = fr.w; s.h = fr.h; s.rot = 0; s.name = fresh.name;
        s.tx = Object.assign(fresh.tx, { ps: keep.ps });
        if (keep.vert && !fr.vert) delete s.tx.vert;
        out.push(s);
      } else if (s) {
        s.ph = { type: fr.type, idx: fr.idx || 0 };
        if (s.type === 'image') { const ar = s.w / s.h, fa = fr.w / fr.h; if (ar > fa) { s.w = fr.w; s.h = fr.w / ar; } else { s.h = fr.h; s.w = fr.h * ar; } s.x = fr.x + (fr.w - s.w) / 2; s.y = fr.y + (fr.h - s.h) / 2; }
        else { s.x = fr.x; s.y = fr.y; s.w = fr.w; if (s.type !== 'table') s.h = fr.h; }
        out.push(s);
      } else out.push(M.makePlaceholder(fr, design));
    }
    for (const s of old) {
      if (used.has(s)) continue;
      if (s.ph && s.type === 'text' && T.isEmpty(s.tx) && !['dt', 'ftr', 'sldNum'].includes(s.ph.type)) continue; /* drop empty placeholders */
      if (s.ph && !['dt', 'ftr', 'sldNum'].includes(s.ph.type)) { delete s.ph; s.name = (s.name || 'Text').replace(/Placeholder|Title|Subtitle/g, 'Text').trim(); }
      out.push(s);
    }
    slide.layout = key;
    slide.shapes = out;
  };

  /** Apply a design template to slides (or all). Returns new design id. */
  M.applyDesign = function (pres, key, slides, opts) {
    const d = M.buildDesign(key, pres.W, pres.H);
    if (opts && opts.keepColors) d.colors = opts.keepColors;
    pres.designs[d.id] = d;
    const targets = slides || pres.slides;
    for (const s of targets) {
      const old = pres.designs[s.design];
      s.design = d.id;
      /* move untouched placeholders to the new master frames */
      const frames = M.layoutFrames(s.layout, d);
      const oldFrames = old ? M.layoutFrames(s.layout, old) : [];
      for (const sh of s.shapes) {
        if (!sh.ph) continue;
        const of = oldFrames.find((f) => f.type === sh.ph.type && (f.idx || 0) === (sh.ph.idx || 0));
        const nf = frames.find((f) => f.type === sh.ph.type && (f.idx || 0) === (sh.ph.idx || 0));
        if (nf && (!of || (Math.abs(of.x - sh.x) < 1 && Math.abs(of.y - sh.y) < 1 && Math.abs(of.w - sh.w) < 1))) {
          sh.x = nf.x; sh.y = nf.y; sh.w = nf.w; if (sh.type !== 'table') sh.h = nf.h;
          if (sh.ph.type === 'ctrTitle') sh.tx.lst = Object.assign({}, sh.tx.lst, { 0: Object.assign({}, (sh.tx.lst || {})[0], { algn: d.ctrTitleAlgn }) });
          if (sh.ph.type === 'subTitle') sh.tx.lst = Object.assign({}, sh.tx.lst, { 0: Object.assign({}, (sh.tx.lst || {})[0], { algn: d.subTitleAlgn }) });
        }
      }
    }
    M.gcDesigns(pres);
    return d.id;
  };
  M.gcDesigns = function (pres) {
    const used = new Set(pres.slides.map((s) => s.design));
    const ids = Object.keys(pres.designs);
    for (const id of ids) if (!used.has(id) && !pres.designs[id].keep?.master && ids.length > 1 && used.size) delete pres.designs[id];
  };

  /* ---------- shape factories ---------- */
  M.newShape = function (pres, slide, geom, x, y, w, h) {
    const d = M.design(pres, slide);
    const def = pres.shapeDefaults || {};
    const sh = {
      id: L.uid('s'), type: 'shape', name: (L.geom.get(geom).label || 'AutoShape') + ' ' + (slide.shapes.length + 1),
      geom, x, y, w, h, rot: 0,
      fill: def.fill ? L.clone(def.fill) : { t: 'solid', c: 'accent1', a: 1 },
      line: def.line ? L.clone(def.line) : { c: 'tx1', w: 0.75, dash: 'solid' },
      shadow: def.shadow ? L.clone(def.shadow) : null,
      tx: T.body([T.para('', { algn: 'ctr' })], { anchor: 'ctr', wrap: true }),
    };
    const g = L.geom.get(geom);
    if (g.noFill || g.line) sh.fill = { t: 'none' };
    if (g.adj) sh.adj = L.clone(g.adj);
    if (g.action) sh.link = { action: { home: 'first', back: 'prev', next: 'next', begin: 'first', end: 'last', return: 'lastViewed' }[g.action] || null };
    if (sh.link && !sh.link.action) delete sh.link;
    void d;
    return sh;
  };
  M.newLine = function (pres, slide, geom, p1, p2, arrows) {
    const def = pres.lineDefaults || {};
    const sh = { id: L.uid('s'), type: 'line', name: 'Line ' + (slide.shapes.length + 1), geom: geom || 'line', x: 0, y: 0, w: 0, h: 0, rot: 0, line: def.line ? L.clone(def.line) : { c: 'tx1', w: 0.75, dash: 'solid' } };
    L.geom.setLineEnds(sh, p1, p2);
    if (arrows === 'end' || arrows === 'both') sh.line.tail = { type: 'triangle', w: 'med', len: 'med' };
    if (arrows === 'both') sh.line.head = { type: 'triangle', w: 'med', len: 'med' };
    return sh;
  };
  M.newTextBox = function (pres, slide, x, y, w, h, wrap) {
    return {
      id: L.uid('s'), type: 'text', name: 'Text Box ' + (slide.shapes.length + 1), geom: 'rect', x, y, w, h, rot: 0,
      fill: { t: 'none' }, line: { t: 'none' },
      tx: T.body([T.para('')], { anchor: 't', wrap: wrap !== false, autofit: 'shape' }),
    };
  };
  M.newImage = function (slide, media, x, y, w, h) {
    return { id: L.uid('s'), type: 'image', name: 'Picture ' + (slide.shapes.length + 1), geom: 'rect', x, y, w, h, rot: 0, media, crop: { l: 0, t: 0, r: 0, b: 0 }, line: { t: 'none' }, lockAspect: true, img: {} };
  };
  M.newTable = function (pres, slide, rows, cols, x, y, w) {
    const d = M.design(pres, slide);
    const cw = w / cols, rh = 0;
    void d;
    const border = { c: 'tx1', w: 1, dash: 'solid' };
    const tbl = { cols: Array(cols).fill(L.round(cw, 2)), rows: [], firstRow: true, bandRow: false };
    for (let r = 0; r < rows; r++) {
      const cells = [];
      for (let c = 0; c < cols; c++) cells.push({ tx: T.body([T.para('', {}, { sz: 18 })], { anchor: 't', ins: [7.2, 3.6, 7.2, 3.6] }), fill: { t: 'none' }, bd: { l: L.clone(border), r: L.clone(border), t: L.clone(border), b: L.clone(border) } });
      tbl.rows.push({ h: Math.max(rh, 28.8), cells });
    }
    return { id: L.uid('s'), type: 'table', name: 'Table ' + (slide.shapes.length + 1), x, y, w, h: rows * 28.8, rot: 0, tbl };
  };

  /* ---------- shape lookup ---------- */
  M.walk = function (shapes, fn, parent) {
    for (const s of shapes) {
      if (fn(s, parent) === false) return false;
      if (s.type === 'group' && s.kids) if (M.walk(s.kids, fn, s) === false) return false;
    }
    return true;
  };
  M.find = function (slide, id) {
    let hit = null, par = null;
    M.walk(slide.shapes, (s, p) => { if (s.id === id) { hit = s; par = p; return false; } return true; });
    return hit ? { shape: hit, parent: par, list: par ? par.kids : slide.shapes } : null;
  };
  M.shapeById = (slide, id) => { const f = M.find(slide, id); return f ? f.shape : null; };
  M.slideTitle = function (slide) {
    const t = slide.shapes.find((s) => s.ph && (s.ph.type === 'title' || s.ph.type === 'ctrTitle'));
    return t && t.tx ? T.plain(t.tx).replace(/\n/g, ' ').trim() : '';
  };
  M.groupBounds = function (g) {
    return L.unionBounds(g.kids.map((k) => (k.type === 'group' ? M.groupBounds(k) : L.rotBounds(k))));
  };
  /* move a shape (and group children) */
  // A retained graphic frame has no Office-supported rotation transform. Keep its
  // preview and payload together until an editor for that object owns the transform.
  M.canRotate = s => !['table', 'chart'].includes(s.type) && !s.keep?.frame && (s.kids || []).every(M.canRotate);
  M.translate = function (s, dx, dy) {
    s.x += dx; s.y += dy;
    if (s.type === 'group') for (const k of s.kids) M.translate(k, dx, dy);
  };
  /* scale a group's children from old box to new box */
  M.scaleGroup = function (g, ob, nb) {
    const sx = ob.w ? nb.w / ob.w : 1, sy = ob.h ? nb.h / ob.h : 1;
    const each = (k) => {
      k.x = nb.x + (k.x - ob.x) * sx; k.y = nb.y + (k.y - ob.y) * sy; k.w *= sx; k.h *= sy;
      if (k.type === 'group') for (const c of k.kids) each(c);
    };
    for (const k of g.kids) each(k);
  };
  /* Deep copy with fresh ids */
  M.dupMany = function (shapes) {
    const copies = L.opc.duplicate(shapes);
    const re = (x) => { x.id = L.uid('s'); if (x.kids) x.kids.forEach(re); };
    for (const c of copies) {
      re(c);
      if (c.ph && !['dt', 'ftr', 'sldNum'].includes(c.ph.type)) delete c.ph;
    }
    return copies;
  };
  M.dup = s => M.dupMany([s])[0];
  M.dupSlide = function (slide) {
    const c = L.opc.duplicate(slide);
    c.id = L.uid('sl');
    if (c.keep) c.keep.copy = c.id;
    const map = {};
    const re = (x) => {
      const n = L.uid('s'); map[x.id] = n;
      if (x.keep?.commentAnchor?.owner === x.id) x.keep.commentAnchor.owner = n;
      x.id = n; if (x.kids) x.kids.forEach(re);
    };
    c.shapes.forEach(re);
    c.anims = (c.anims || []).map((a) => Object.assign({}, a, { id: L.uid('a'), sid: map[a.sid] || a.sid }));
    if (c.keep?.anims) c.keep.anims = L.preserve.anims(c);
    return c;
  };
  M.allText = function (slide) {
    const out = [];
    M.walk(slide.shapes, (s) => {
      if (s.tx) out.push(T.plain(s.tx));
      if (s.type === 'table') for (const r of s.tbl.rows) for (const c of r.cells) out.push(T.plain(c.tx));
      if (s.type === 'wordart') out.push(s.wa.text);
    });
    return out.join('\n');
  };

  /* ---------- undo / redo ---------- */
  const H = (L.hist = { undo: [], redo: [], limit: 100, dirty: false, lastLabel: '' });
  H.snapshot = () => JSON.stringify({ W: L.pres.W, H: L.pres.H, firstNum: L.pres.firstNum, designs: L.pres.designs, slides: L.pres.slides, show: L.pres.show, hf: L.pres.hf, props: L.pres.props, title: L.pres.title, keep: L.pres.keep, ooxmlFormat: L.pres.ooxmlFormat, losses: L.pres.losses || [], shapeDefaults: L.pres.shapeDefaults, lineDefaults: L.pres.lineDefaults });
  H.restore = (snap) => { Object.assign(L.pres, JSON.parse(snap)); };
  /** Record state before a mutation. coalesceKey merges rapid repeats (typing). */
  H.push = (label, coalesceKey) => {
    const now = Date.now();
    if (coalesceKey && H._key === coalesceKey && now - H._t < 1200) { H._t = now; H.dirty = true; return; }
    H._key = coalesceKey || null; H._t = now;
    H.undo.push({ label: label || 'Edit', snap: H.snapshot(), sel: L.ed ? L.ed.saveSel() : null });
    if (H.undo.length > H.limit) H.undo.shift();
    H.redo.length = 0;
    H.dirty = true;
    H.lastLabel = label || 'Edit';
  };
  H.breakCoalesce = () => { H._key = null; };
  H.doUndo = () => {
    const e = H.undo.pop();
    if (!e) return false;
    H.redo.push({ label: e.label, snap: H.snapshot(), sel: L.ed ? L.ed.saveSel() : null });
    H.restore(e.snap);
    H._key = null;
    L.bus.emit('history', e);
    return true;
  };
  H.doRedo = () => {
    const e = H.redo.pop();
    if (!e) return false;
    H.undo.push({ label: e.label, snap: H.snapshot(), sel: L.ed ? L.ed.saveSel() : null });
    H.restore(e.snap);
    H._key = null;
    L.bus.emit('history', e);
    return true;
  };
  H.clear = () => { H.undo.length = 0; H.redo.length = 0; H.dirty = false; H._key = null; };
})();
