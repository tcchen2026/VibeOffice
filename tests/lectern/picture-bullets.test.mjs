// A picture bullet (a:buBlip) is drawn as its picture and written back as one, with the same image, through a save
// and a reopen; it used to be read as a square character, and a paragraph written from the model lost the picture.
// node --test tests/lectern/picture-bullets.test.mjs
import { after, before, test } from 'node:test';
import assert from 'node:assert/strict';
import { ensureServer, openPage } from '../../tools/ooxml/cdp.mjs';

let page, stop;
before(async () => {
  stop = await ensureServer(process.env.BASE);
  page = await openPage('lectern', { persistence: false, automated: true });
});
after(async () => {
  try { await page?.close(); } finally { stop?.(); }
});

test('a picture bullet is drawn as its picture, saved as a:buBlip with its image, and read back as one', async () => {
  const out = await page.evaluate(`(async () => {
    const c = document.createElement('canvas'); c.width = c.height = 8;
    const g = c.getContext('2d'); g.fillStyle = '#1060E0'; g.fillRect(0, 0, 8, 8);
    const png = await new Promise((r) => c.toBlob(r, 'image/png')), bytes = new Uint8Array(await png.arrayBuffer());
    const sh = L.pres.slides[0].shapes.find((s) => s.tx);
    sh.tx.ps = [L.txt.para('Picture bullet')]; sh.tx.ps[0].pp = Object.assign({}, sh.tx.ps[0].pp, { bu: { t: 'blip', media: L.media.add(png, 'bullet.png') } });
    L.ed.render(); await new Promise((r) => setTimeout(r, 300));
    const p = document.querySelector('.sl p.pa.bu-img');
    const drawn = !!p && /^url\\("blob:/.test(p.style.getPropertyValue('--bu-img'));
    const saved = await L.pptx.write(L.pres), pk = await L.opc.open(new Uint8Array(await saved.arrayBuffer()));
    const xml = pk.text('ppt/slides/slide1.xml'), rid = (/<a:buBlip><a:blip r:embed="([^"]+)"/.exec(xml) || [])[1];
    const rel = rid && pk.rels('ppt/slides/slide1.xml').find((r) => r.id === rid), image = rel && pk.bytes(rel.part);
    await L.app.openFile(new File([saved], 'bullets.pptx')); await new Promise((r) => setTimeout(r, 500));
    const bu = L.pres.slides[0].shapes.flatMap((s) => (s.tx ? s.tx.ps : [])).map((q) => q.pp && q.pp.bu).find(Boolean);
    const again = bu && bu.t === 'blip' ? new Uint8Array(await L.media.get(bu.media).blob.arrayBuffer()) : null;
    return { drawn, written: !!rid, sameImage: !!image && image.length === bytes.length && image.every((b, i) => b === bytes[i]),
      read: bu && bu.t, reread: !!again && again.length === bytes.length };
  })()`);
  assert.equal(out.drawn, true, 'drawn as the picture');
  assert.equal(out.written, true, 'saved as a:buBlip');
  assert.equal(out.sameImage, true, 'with its image');
  assert.equal(out.read, 'blip', 'read back as a picture bullet');
  assert.equal(out.reread, true);
});
