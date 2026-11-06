// Word 2010's text effects on runs (w14:textFill, textOutline, shadow, glow) are drawn, against Word 2016's prints of
// LibreOffice's TextEffects_TextFill.docx, TextEffects_TextOutline.docx and semi-transparent-text.docx (2026-11-07):
// a solid fill with its transparency (w14's alpha is a transparency, DrawingML's an opacity: 74000 is a quarter
// opaque), a gradient through the letters, the outline and shadow; Quire drew such text black or in its plain colour
// (a library's cover title, orange with a shadow in Word, was black). A colour set in Quire drops the fill, as the save
// does (docs/LOSSLESS_SAVE.md: the saved w14:textFill belongs to the colour).
// node --test tests/quire/text-effects.test.mjs
import { after, before, test } from 'node:test';
import assert from 'node:assert/strict';
import { ensureServer, openPage } from '../../tools/ooxml/cdp.mjs';

let page, stop;
before(async () => {
  stop = await ensureServer(process.env.BASE);
  page = await openPage('quire', { persistence: false, automated: true });
});
after(async () => {
  try { await page?.close(); } finally { stop?.(); }
});

const W14 = 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
const run = (text, rPr) => '<w:r><w:rPr>' + rPr + '</w:rPr><w:t>' + text + '</w:t></w:r>';
const BODY = '<w:p>' +
  run('Faint', '<w:color w:val="FF0000"/><w14:textFill><w14:solidFill><w14:srgbClr w14:val="FF0000"><w14:alpha w14:val="74000"/></w14:srgbClr></w14:solidFill></w14:textFill>') +
  run('Grad', '<w14:textFill><w14:gradFill><w14:gsLst><w14:gs w14:pos="0"><w14:srgbClr w14:val="FF0000"/></w14:gs><w14:gs w14:pos="100000"><w14:srgbClr w14:val="0000FF"/></w14:gs></w14:gsLst><w14:lin w14:ang="0" w14:scaled="0"/></w14:gradFill></w14:textFill>') +
  run('Edge', '<w14:textOutline w14:w="9525" w14:cap="flat" w14:cmpd="sng" w14:algn="ctr"><w14:solidFill><w14:srgbClr w14:val="00B050"/></w14:solidFill></w14:textOutline>') +
  run('Shade', '<w14:shadow w14:blurRad="38100" w14:dist="19050" w14:dir="2700000" w14:sx="100000" w14:sy="100000" w14:kx="0" w14:ky="0" w14:algn="tl"><w14:srgbClr w14:val="000000"><w14:alpha w14:val="60000"/></w14:srgbClr></w14:shadow>') +
  '</w:p>';

// each run's drawn colour, background clip, stroke and shadow, by its text
async function styles() {
  return page.evaluate(`(() => {
    const out = {};
    for (const s of document.querySelectorAll('#pages .pc span')) {
      const t = s.textContent.trim(); if (!['Faint', 'Grad', 'Edge', 'Shade'].includes(t)) continue;
      const cs = getComputedStyle(s);
      out[t] = { color: cs.color, clip: cs.webkitBackgroundClip || cs.backgroundClip, stroke: cs.webkitTextStrokeWidth + ' ' + cs.webkitTextStrokeColor, shadow: cs.textShadow };
    }
    return out;
  })()`);
}

test('text effects are drawn: fill with its transparency, gradient, outline and shadow', async () => {
  await page.evaluate(`(async () => {
    const N = L.opc.NS, part = 'word/document.xml';
    const files = [
      { name: '[Content_Types].xml', data: '<Types xmlns="' + N.ct + '"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/' + part + '" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>' },
      { name: '_rels/.rels', data: '<Relationships xmlns="' + N.pkg + '"><Relationship Id="rId1" Type="' + N.rel + '/officeDocument" Target="' + part + '"/></Relationships>' },
      { name: part, data: '<w:document xmlns:w="' + N.w + '" ${W14}><w:body>' + ${JSON.stringify(BODY)} + '<w:sectPr/></w:body></w:document>' },
    ];
    await L.app.openFile(new File([await L.zip.write(files)], 'effects.docx'));
    await new Promise((r) => setTimeout(r, 300));
  })()`);
  assert.deepEqual(page.errors, []);
  const s = await styles();
  assert.equal(s.Faint.color, 'rgba(255, 0, 0, 0.26)');
  assert.equal(s.Grad.clip, 'text');
  assert.equal(s.Grad.color, 'rgba(0, 0, 0, 0)');
  assert.match(s.Edge.stroke, /^1(\.\d+)?px rgb\(0, 176, 80\)$/);
  assert.match(s.Shade.shadow, /^rgba\(0, 0, 0, 0\.4\) /);
  // a colour set here drops the fill, as Word and the save do
  await page.evaluate(`(async () => {
    const p = L.D.doc.main.blocks[0], it = p.runs.find((x) => x.t === 'text' && x.text === 'Faint');
    it.rPr = Object.assign({}, it.rPr, { color: '0000FF' });
    L.R.clearCaches(); L.layout.render();
    await new Promise((r) => setTimeout(r, 300));
  })()`);
  assert.equal((await styles()).Faint.color, 'rgb(0, 0, 255)');
});
