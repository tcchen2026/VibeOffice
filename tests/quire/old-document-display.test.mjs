// How Quire shows what Word 97-2003 documents carry (I24 E3 findings, checked against Word's own prints): a
// picture's borders set per side, a frame with no vertical position, and text in the Symbol font.
// node --test tests/quire/old-document-display.test.mjs
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

// opens a document whose body is `body` (with a 1×1 picture as rImg) and returns probe(result)
async function open(body, probe) {
  const r = await page.evaluate(`(async () => {
    const K = L.opc, N = K.NS, part = 'word/document.xml';
    const c = document.createElement('canvas'); c.width = c.height = 4; c.getContext('2d').fillRect(0, 0, 4, 4);
    const png = new Uint8Array(await (await new Promise((r) => c.toBlob(r, 'image/png'))).arrayBuffer());
    const ns = 'xmlns:w="' + N.w + '" xmlns:r="' + N.rel + '" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w10="urn:schemas-microsoft-com:office:word"';
    const files = [
      { name: '[Content_Types].xml', data: '<Types xmlns="' + N.ct + '"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/' + part + '" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>' },
      { name: '_rels/.rels', data: '<Relationships xmlns="' + N.pkg + '"><Relationship Id="rId1" Type="' + N.rel + '/officeDocument" Target="' + part + '"/></Relationships>' },
      { name: 'word/_rels/document.xml.rels', data: '<Relationships xmlns="' + N.pkg + '"><Relationship Id="rImg" Type="' + N.rel + '/image" Target="media/p.png"/></Relationships>' },
      { name: part, data: '<w:document ' + ns + '><w:body>' + ${JSON.stringify(body)} + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>' },
      { name: 'word/media/p.png', data: png },
    ];
    await L.app.openFile(new File([await L.zip.write(files)], 'old.docx'));
    L.layout.setZoom(1);
    await new Promise((r) => setTimeout(r, 300));
    return (${probe})();
  })()`);
  assert.deepEqual(page.errors, []);
  return r;
}

test('a picture\'s borders set per side are drawn, each in its own style and colour', async () => {
  const r = await open('<w:p><w:r><w:pict><v:rect style="width:40pt;height:40pt" stroked="f" o:bordertopcolor="#FF0000" o:borderleftcolor="black" o:borderbottomcolor="black" o:borderrightcolor="black">' +
    '<v:imagedata r:id="rImg" o:title=""/><w10:bordertop type="dash" width="12"/><w10:borderleft type="single" width="18"/><w10:borderbottom type="dot" width="24"/><w10:borderright type="double" width="12"/></v:rect></w:pict></w:r></w:p>',
  `() => { const s = getComputedStyle(document.querySelector('#pages .img > .img-sides')); return [s.borderTopStyle, s.borderTopColor, s.borderLeftStyle, s.borderBottomStyle, s.borderRightStyle]; }`);
  assert.deepEqual(r, ['dashed', 'rgb(255, 0, 0)', 'solid', 'dotted', 'double']);
});

test('a frame with no vertical position stands beside its own paragraph, not at the top of the margin', async () => {
  const para = (t) => '<w:p><w:r><w:t>' + t + '</w:t></w:r></w:p>';
  const r = await open(para('First paragraph.').repeat(12) + '<w:p><w:pPr><w:framePr w:hSpace="180" w:vSpace="180" w:vAnchor="margin" w:hAnchor="margin" w:xAlign="right"/></w:pPr>' +
    '<w:r><w:pict><v:shape style="width:100pt;height:60pt" type="#_x0000_t75"><v:imagedata r:id="rImg" o:title=""/></v:shape></w:pict></w:r></w:p>' + para('Text beside the frame.'),
  `() => { const img = document.querySelector('#pages span.img').getBoundingClientRect(), paras = [...document.querySelectorAll('#pages span')].filter((s) => s.textContent === 'First paragraph.');
    return { img: img.top, lastPara: paras[paras.length - 1].getBoundingClientRect().bottom, firstPara: paras[0].getBoundingClientRect().top }; }`);
  assert.ok(r.img >= r.lastPara - 1, `the frame (top ${r.img}) is below the twelve paragraphs before it (bottom ${r.lastPara})`);
});

test('text in the Symbol font shows its Greek letters, typed or in Word\'s private-use form', async () => {
  const sym = (t) => '<w:r><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/></w:rPr><w:t xml:space="preserve">' + t + '</w:t></w:r>';
  const r = await open('<w:p><w:r><w:t xml:space="preserve">angle </w:t></w:r>' + sym('q') + '<w:r><w:t xml:space="preserve">, </w:t></w:r>' + sym('') + sym('abg') + '</w:p>',
    `() => document.querySelector('#pages').textContent.replace(/\\s+/g, ' ').trim()`);
  assert.match(r, /angle θ, θαβγ/);
});
