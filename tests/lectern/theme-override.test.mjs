// A slide with colours of its own (a theme override, as PowerPoint writes for a slide whose colour scheme differs
// from its master's, and for every such slide of a 97-2003 file it converts) is drawn in them; its master keeps its
// own. Such slides used to take their master's colours. node --test tests/lectern/theme-override.test.mjs
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

test('a slide\'s theme override gives its colours; the master keeps its own, and a save keeps the override', async () => {
  const out = await page.evaluate(`(async () => {
    const A = 'http://schemas.openxmlformats.org/drawingml/2006/main', R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    const slot = (n, v) => '<a:' + n + '><a:srgbClr val="' + v + '"/></a:' + n + '>';
    const scheme = [['dk1', '000000'], ['lt1', 'C04020'], ['dk2', '1F497D'], ['lt2', 'EEECE1'], ['accent1', '4F81BD'], ['accent2', 'C0504D'], ['accent3', '9BBB59'],
      ['accent4', '8064A2'], ['accent5', '4BACC6'], ['accent6', 'F79646'], ['hlink', '0000FF'], ['folHlink', '800080']].map(([n, v]) => slot(n, v)).join('');
    const blob = await L.pptx.write(L.pres), zip = await L.zip.read(new Uint8Array(await blob.arrayBuffer())), parts = [];
    for (const [name, e] of zip) {
      let data = await e.bytes();
      const text = () => new TextDecoder().decode(data);
      if (name === '[Content_Types].xml') data = text().replace('</Types>', '<Override PartName="/ppt/theme/themeOverride1.xml" ContentType="application/vnd.openxmlformats-officedocument.themeOverride+xml"/></Types>');
      if (name === 'ppt/slides/_rels/slide1.xml.rels') data = text().replace('</Relationships>', '<Relationship Id="rId99" Type="' + R + '/themeOverride" Target="../theme/themeOverride1.xml"/></Relationships>');
      parts.push({ name, data });
    }
    parts.push({ name: 'ppt/theme/themeOverride1.xml', data: '<a:themeOverride xmlns:a="' + A + '"><a:clrScheme name="Red">' + scheme + '</a:clrScheme></a:themeOverride>' });
    await L.app.openFile(new File([await L.zip.write(parts)], 'override.pptx'));
    await new Promise((r) => setTimeout(r, 500));
    const M = L.model, s = L.pres.slides[0];
    const saved = await L.opc.open(new Uint8Array(await (await L.pptx.write(L.pres)).arrayBuffer()));
    return { slide: M.resolveColor('bg1', M.design(L.pres, s)), master: M.resolveColor('bg1', L.pres.designs[s.design]),
      kept: /themeOverride/.test(saved.text('ppt/slides/_rels/slide1.xml.rels') || '') };
  })()`);
  assert.equal(out.slide, '#C04020', 'the slide\'s own background colour');
  assert.equal(out.master, '#FFFFFF', 'the master\'s, unchanged');
  assert.equal(out.kept, true, 'the save keeps the slide\'s override');
});
