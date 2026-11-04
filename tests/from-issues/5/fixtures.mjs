// Browser-side synthetic controls for column painting, independent of the issue's private document text.
export async function makeColumnsFixture(kind) {
  const { D } = L, doc = D.newDoc({ modern: true });
  Object.assign(doc.sect, { pgW: 612, pgH: 792, mt: 72, mb: 72, ml: 72, mr: 72, gutter: 0, docGrid: null });
  doc.sect.cols = kind === 'single' ? { n: 1 } : kind === 'unequal'
    ? { n: 3, eq: false, w: [{ w: 120, space: 18 }, { w: 140, space: 30 }, { w: 160, space: 0 }] }
    : { n: 2, eq: true, space: 24 };
  const canvas = document.createElement('canvas'); canvas.width = canvas.height = 16;
  canvas.getContext('2d').fillRect(0, 0, 16, 16);
  const blob = await new Promise(resolve => canvas.toBlob(resolve));
  const media = L.media.add(blob, 'column-control.png');
  doc.main.blocks = [];
  for (let i = 0; i < doc.sect.cols.n; i++) {
    if (i) doc.main.blocks.push(D.para([D.item('br', { type: 'column' })]));
    const drawing = D.item(kind === 'shapes' ? 'shape' : 'img', { media, w: 320, h: 60, alt: 'Column ' + (i + 1) });
    if (kind === 'rotated') Object.assign(drawing, { rot: 30, flipH: true });
    if (kind === 'shapes') Object.assign(drawing, { geom: 'rect', fill: { t: 'solid', c: '#000000' }, line: { t: 'none' } });
    doc.main.blocks.push(D.para([drawing], { ind: { l: i ? -60 : -30 }, sp: { a: 0, b: 0 } }));
    if (!i && kind === 'floating') {
      for (const [wrap, y] of [['none', 180], ['square', 300]]) {
        doc.main.blocks.push(D.para([D.item('img', { media, w: 300, h: 60, alt: 'Floating ' + wrap,
          float: { wrap, posH: { rel: 'page', off: 160 }, posV: { rel: 'page', off: y }, layoutInCell: true, allowOverlap: true, dist: { l: 0, r: 0, t: 0, b: 0 } } })]));
      }
    }
    doc.main.blocks.push(D.para([D.text('Text in column ' + (i + 1))]));
  }
  return doc;
}
