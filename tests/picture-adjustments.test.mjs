// A picture's brightness and contrast are drawn as Office draws them (L.lumFilter): a grey ramp of 11 bands under
// VML and DrawingML settings, compared band by band with the greys Word 2016 and PowerPoint 2016 printed for the
// same settings (tmp/lum, 2026-11-06). The browser's multiplying brightness() drew Washout's black as 89, not 205.
// node --test tests/picture-adjustments.test.mjs
import { after, test } from 'node:test';
import assert from 'node:assert/strict';
import { ensureServer, openPage } from '../tools/ooxml/cdp.mjs';

let stop;
after(() => stop?.());

// Office's output grey for input greys 0, 25.5, …, 255
const OFFICE = {
  none: [0, 26, 51, 76, 102, 128, 153, 178, 204, 230, 255],
  'vml washout': [205, 213, 221, 228, 236, 244, 251, 255, 255, 255, 255],
  'vml gain 2': [0, 0, 0, 25, 77, 129, 179, 229, 255, 255, 255],
  'vml blacklevel -0.25': [0, 0, 0, 0, 0, 1, 26, 51, 77, 103, 128],
  'vml gain 0.5 blacklevel 0.2': [140, 153, 166, 178, 191, 204, 217, 229, 242, 255, 255],
  'bright 40%': [102, 128, 153, 178, 204, 230, 255, 255, 255, 255, 255],
  'contrast -40%': [51, 66, 81, 96, 112, 127, 142, 157, 173, 189, 204],
  'contrast 40%': [0, 0, 0, 41, 84, 128, 169, 211, 254, 255, 255],
  'bright 20% contrast 30%': [7, 44, 80, 115, 153, 190, 225, 255, 255, 255, 255],
};

// the page draws the ramp, places the pictures and returns their rectangles; then a screenshot is sampled
async function drawn(app, setup) {
  stop = stop || await ensureServer(process.env.BASE);
  const page = await openPage(app, { persistence: false, automated: true });
  try {
    await page.send('Emulation.setDeviceMetricsOverride', { width: 1400, height: 1600, deviceScaleFactor: 1, mobile: false });
    const rects = await page.evaluate(`(async () => {
      const c = document.createElement('canvas'); c.width = 660; c.height = 20;
      const g = c.getContext('2d');
      for (let k = 0; k < 11; k++) { const v = Math.round(255 * k / 10); g.fillStyle = 'rgb(' + v + ',' + v + ',' + v + ')'; g.fillRect(k * 60, 0, 60, 20); }
      const png = new Uint8Array(await (await new Promise((r) => c.toBlob(r, 'image/png'))).arrayBuffer());
      return (${setup})(png);
    })()`);
    await new Promise((r) => setTimeout(r, 500));
    const shot = await page.send('Page.captureScreenshot', { format: 'png' });
    const greys = await page.evaluate(`(async () => {
      const im = new Image(); im.src = 'data:image/png;base64,${shot.data}'; await im.decode();
      const c = document.createElement('canvas'); c.width = im.width; c.height = im.height;
      const g = c.getContext('2d'); g.drawImage(im, 0, 0);
      return ${JSON.stringify(rects)}.map(([x, y, w, h]) => Array.from({ length: 11 }, (_, k) => g.getImageData(Math.round(x + (k + 0.5) * w / 11), Math.round(y + h / 2), 1, 1).data[0]));
    })()`);
    assert.deepEqual(page.errors, []);
    return greys;
  } finally { await page.close(); }
}

const close = (got, want, name) => assert.ok(got.every((v, k) => Math.abs(v - want[k]) <= 2), `${name}: drew ${got}, Office ${want}`);

test('Quire draws VML and DrawingML picture adjustments as Word does', async () => {
  const cases = [['none', ''], ['vml washout', 'gain="19661f" blacklevel="22938f"'], ['vml gain 2', 'gain="2"'], ['vml blacklevel -0.25', 'blacklevel="-0.25"'],
    ['vml gain 0.5 blacklevel 0.2', 'gain="0.5" blacklevel="0.2"'], ['bright 40%', 'bright="40000"'], ['contrast -40%', 'contrast="-40000"'],
    ['contrast 40%', 'contrast="40000"'], ['bright 20% contrast 30%', 'bright="20000" contrast="30000"']];
  const greys = await drawn('quire', `async (png) => {
    const cases = ${JSON.stringify(cases)}, N = L.opc.NS, part = 'word/document.xml';
    const pic = ([name, a], i) => name.startsWith('vml') || name === 'none'
      ? '<w:p><w:r><w:pict><v:shape id="p' + i + '" type="#_x0000_t75" style="width:432pt;height:14.4pt"><v:imagedata r:id="rImg" o:title="" ' + a + '/></v:shape></w:pict></w:r></w:p>'
      : '<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="5486400" cy="182880"/><wp:docPr id="' + (i + 1) + '" name="Picture ' + (i + 1) + '"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="' + (i + 1) + '" name="ramp.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rImg"><a:lum ' + a + '/></a:blip><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="5486400" cy="182880"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>';
    const ns = 'xmlns:w="' + N.w + '" xmlns:r="' + N.rel + '" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="' + N.a + '" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';
    const files = [
      { name: '[Content_Types].xml', data: '<Types xmlns="' + N.ct + '"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/' + part + '" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>' },
      { name: '_rels/.rels', data: '<Relationships xmlns="' + N.pkg + '"><Relationship Id="rId1" Type="' + N.rel + '/officeDocument" Target="' + part + '"/></Relationships>' },
      { name: 'word/_rels/document.xml.rels', data: '<Relationships xmlns="' + N.pkg + '"><Relationship Id="rImg" Type="' + N.rel + '/image" Target="media/ramp.png"/></Relationships>' },
      { name: part, data: '<w:document ' + ns + '><w:body>' + cases.map(pic).join('') + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="360" w:footer="360" w:gutter="0"/></w:sectPr></w:body></w:document>' },
      { name: 'word/media/ramp.png', data: png },
    ];
    await L.app.openFile(new File([await L.zip.write(files)], 'ramp.docx'));
    L.layout.setZoom(1);
    L.layout.scroller.scrollTop = 0;
    await new Promise((r) => setTimeout(r, 300));
    return [...document.querySelectorAll('#pages span.img')].map((e) => { const r = e.getBoundingClientRect(); return [r.left, r.top, r.width, r.height]; });
  }`);
  assert.equal(greys.length, cases.length);
  cases.forEach(([name], i) => close(greys[i], OFFICE[name], name));
});

test('Lectern draws picture brightness and contrast as PowerPoint does', async () => {
  const cases = [['none'], ['bright 40%', 0.4, 0], ['contrast -40%', 0, -0.4], ['contrast 40%', 0, 0.4], ['bright 20% contrast 30%', 0.2, 0.3], ['vml washout', 'wash']];
  const greys = await drawn('lectern', `async (png) => {
    const cases = ${JSON.stringify(cases)};
    await L.app.insertPictureFiles(cases.map((c, i) => new File([png], 'ramp' + i + '.png', { type: 'image/png' })));
    const E = L.ed;
    const shapes = E.slide().shapes.filter((s) => s.img || s.media).slice(-cases.length);
    shapes.forEach((sh, i) => {
      const [, b, c] = cases[i];
      sh.x = L.pres.W * 0.05; sh.y = L.pres.H * (0.04 + i * 0.15); sh.w = L.pres.W * 0.9; sh.h = L.pres.H * 0.08;
      sh.img = Object.assign({}, sh.img, b === 'wash' ? { mode: 'wash' } : { bright: b || 0, contrast: c || 0 });
    });
    E.sel = [];
    E.render();
    await new Promise((r) => setTimeout(r, 300));
    const ids = shapes.map((s) => s.id);
    // the slide being edited, not its thumbnail: the widest drawing of each picture
    return ids.map((id) => {
      const r = [...document.querySelectorAll('[data-sid="' + id + '"] img')].map((e) => e.getBoundingClientRect()).sort((a, b) => b.width - a.width)[0];
      return [r.left, r.top, r.width, r.height];
    });
  }`);
  cases.forEach(([name], i) => close(greys[i], OFFICE[name], name));
});
