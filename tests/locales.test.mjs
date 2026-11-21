// New documents in the UI's language (docs/i18n.md, Defaults; public/common/locales.js): each app opened blank with
// ?lang=xx saves what that language's Office makes: the page, line grid, fonts and languages of a document, the font
// and paper of a workbook, the text language and East Asian font of a presentation. English keeps the apps' own.
//   node --test tests/locales.test.mjs
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import { ensureServer, openPage } from '../tools/ooxml/cdp.mjs';

const stop = await ensureServer();
after(() => stop && stop());

/** the parts of the blank document app saves with the UI in lang */
async function saved(app, lang, write, parts) {
  const page = await openPage(app, { lang, persistence: false, automated: true });
  try {
    return await page.evaluate(`(async () => {
      for (let i = 0; i < 100 && !(L.app && ${write.split('(')[0].replace(/^await /, '')}); i++) await new Promise((r) => setTimeout(r, 100));
      const blob = await ${write};
      const zip = await L.zip.read(await blob.arrayBuffer());
      const out = {};
      for (const p of ${JSON.stringify(parts)}) out[p] = zip.get(p) ? await zip.get(p).text() : '';
      return out;
    })()`);
  } finally { await page.close(); }
}
const quire = (lang) => saved('quire', lang, 'L.docx.write(L.D.doc)', ['word/document.xml', 'word/styles.xml', 'word/settings.xml', 'word/theme/theme1.xml']);
const ledger = (lang) => saved('ledger', lang, 'L.xlsxWrite.write(L.app.books[L.app.cur].wb)', ['xl/styles.xml', 'xl/worksheets/sheet1.xml']);
const lectern = (lang) => saved('lectern', lang, 'L.pptx.write(L.pres)', ['ppt/presentation.xml', 'ppt/theme/theme1.xml', 'ppt/slides/slide1.xml']);

for (const [lang, font, pitch, size, tab, charset, row, script, tag] of [
  ['zh-CN', '宋体', 312, 21, 420, 134, 13.5, 'Hans', 'zh-CN'],
  ['zh-Hant', '新細明體', 360, 24, 480, 136, 14.25, 'Hant', 'zh-HK'],   // a tag Word knows: it ignores zh-Hant
]) test('Chinese defaults survive saving in all three apps: ' + lang, async () => {
  const q = await quire(lang), w = q['word/styles.xml'], set = q['word/settings.xml'];
  assert.match(q['word/document.xml'], /<w:pgSz w:w="11906" w:h="16838"/);
  assert.ok(q['word/document.xml'].includes('w:linePitch="' + pitch + '"'));
  assert.ok(w.includes('w:eastAsia="' + font + '"'));
  assert.ok(w.includes('<w:sz w:val="' + size + '"/>'));
  assert.ok(set.includes('<w:defaultTabStop w:val="' + tab + '"/>'));
  assert.ok(set.includes('w:eastAsia="' + tag + '"'));
  const x = await ledger(lang);
  assert.ok(x['xl/styles.xml'].includes('<name val="' + font + '"/>'));
  assert.ok(x['xl/styles.xml'].includes('<charset val="' + charset + '"/>'));
  assert.ok(x['xl/worksheets/sheet1.xml'].includes('defaultRowHeight="' + row + '"'));
  assert.match(x['xl/worksheets/sheet1.xml'], /<pageSetup [^>]*paperSize="9"/);
  const p = await lectern(lang);
  assert.ok(p['ppt/presentation.xml'].includes('lang="' + tag + '"'));
  assert.ok(p['ppt/theme/theme1.xml'].includes('script="' + script + '" typeface="' + font + '"'));
});

test('Quire, Japanese: A4, Office\'s margins, an 18pt line grid, Century with ＭＳ 明朝 10.5pt, headings in ＭＳ ゴシック, punctuation compressed', async () => {
  const x = await quire('ja');
  const doc = x['word/document.xml'], st = x['word/styles.xml'], set = x['word/settings.xml'];
  assert.match(doc, /<w:pgSz w:w="11906" w:h="16838"/);
  assert.match(doc, /<w:pgMar w:top="1985" w:right="1701" w:bottom="1701" w:left="1701" w:header="851" w:footer="992"/);
  assert.match(doc, /<w:docGrid w:type="lines" w:linePitch="360"/);
  const dd = st.match(/<w:docDefaults>.*?<\/w:docDefaults>/s)[0];
  assert.match(dd, /<w:rFonts w:ascii="Century" w:hAnsi="Century" w:eastAsia="ＭＳ 明朝"/);
  assert.match(dd, /<w:sz w:val="21"\/>/);
  assert.match(dd, /<w:lang w:val="en-US" w:eastAsia="ja-JP"\/>/);
  const normal = st.match(/<w:style [^>]*w:styleId="Normal".*?<\/w:style>/s)[0];
  assert.match(normal, /<w:widowControl w:val="(0|false)"\/>/);
  assert.match(normal, /<w:jc w:val="both"\/>/);
  assert.match(set, /<w:defaultTabStop w:val="840"\/>/);
  assert.match(set, /<w:characterSpacingControl w:val="compressPunctuation"\/>/);
  assert.match(set, /<w:themeFontLang w:val="en-US" w:eastAsia="ja-JP"\/>/);
  assert.match(x['word/theme/theme1.xml'], /<a:minorFont>.*<a:font script="Jpan" typeface="ＭＳ 明朝"\/>.*<\/a:minorFont>/s);
  const heading = st.match(/<w:style [^>]*w:styleId="Heading1".*?<\/w:style>/s)[0];
  assert.match(heading, /<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="ＭＳ ゴシック"/, 'headings in the theme\'s heading font, as Japanese Word\'s');
  const plain = st.match(/<w:style [^>]*w:styleId="PlainText".*?<\/w:style>/s);
  if (plain) assert.doesNotMatch(plain[0], /w:eastAsia=/, 'Courier New beside the body\'s ＭＳ 明朝, inherited');
  const zh = (await quire('zh-CN'))['word/styles.xml'].match(/<w:style [^>]*w:styleId="Heading1".*?<\/w:style>/s)[0];
  assert.doesNotMatch(zh, /w:eastAsia=/, 'Chinese headings and text share 宋体');
});

test('Quire, French and German: A4, 2.5 cm margins (German 2 cm at the bottom), the text\'s language', async () => {
  const fr = await quire('fr'), de = await quire('de');
  assert.match(fr['word/document.xml'], /<w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708"/);
  assert.match(de['word/document.xml'], /<w:pgMar w:top="1417" w:right="1417" w:bottom="1134" w:left="1417" w:header="708" w:footer="708"/);
  assert.match(fr['word/styles.xml'], /<w:docDefaults>.*<w:lang w:val="fr-FR"\/>.*<\/w:docDefaults>/s);
  assert.match(de['word/settings.xml'], /<w:themeFontLang w:val="de-DE"\/>/);
  assert.doesNotMatch(fr['word/document.xml'], /<w:docGrid w:type/);
});

test('Quire, English: the app\'s own Letter page, no grid', async () => {
  const x = await quire('en');
  assert.match(x['word/document.xml'], /<w:pgSz w:w="12240" w:h="15840"/);
  assert.doesNotMatch(x['word/document.xml'], /<w:docGrid w:type/);
  assert.match(x['word/settings.xml'], /<w:themeFontLang w:val="en-US"\/>/);
  assert.match(x['word/settings.xml'], /<w:characterSpacingControl w:val="doNotCompress"\/>/);
});

/** what Quire's Font box and status bar show in a new document with the UI in lang: blank, after Latin text,
 *  after East Asian text; with ja, also the fonts of the text after choosing ＭＳ ゴシック, then Arial */
async function shown(lang) {
  const page = await openPage('quire', { lang, persistence: false, automated: true });
  try {
    return await page.evaluate(`(async () => {
      for (let i = 0; i < 100 && !(L.app && L.app.sb && L.ed && L.ed.sel); i++) await new Promise((r) => setTimeout(r, 100));
      const wait = () => new Promise((r) => setTimeout(r, 250)), box = document.getElementById('tb-font');
      /* the box's own refresh: ui.refresh waits for a frame, which a tab in the background doesn't get */
      const st = async () => { box.parentElement._refresh(); L.app.updateStatus(); await wait(); return [box.value, L.app.sb.lang.textContent]; };
      const fonts = () => L.D.doc.main.blocks[0].runs.filter((r) => r.t === 'text').map((r) => [r.rPr.font, r.rPr.fontEA]);
      const out = { blank: await st() };
      L.ed.typeText('abc'); out.latin = await st();
      L.ed.typeText('日本'); out.ea = await st();
      if (${JSON.stringify(lang)} === 'ja') {
        L.ui.exec('selectAll'); L.app.setFont('ＭＳ ゴシック'); out.gothic = fonts();
        L.app.setFont('Arial'); out.arial = fonts();
      }
      return out;
    })()`);
  } finally { await page.close(); }
}

test('Quire, Japanese, Korean, Chinese: the Font box and status bar show the East Asian font and language, as that Word does with its keyboard', async () => {
  const ja = await shown('ja');
  assert.deepEqual([ja.blank, ja.latin, ja.ea], [['ＭＳ 明朝', '日本語'], ['Century', '英語 (米国)'], ['ＭＳ 明朝', '日本語']],
    'blank: the UI\'s language stands in for the keyboard; then the text before the insertion point decides');
  assert.deepEqual(ja.gothic, [['ＭＳ ゴシック', 'ＭＳ ゴシック']], 'an East Asian font is for all the text');
  assert.deepEqual(ja.arial, [['Arial', 'ＭＳ ゴシック']], 'a Western font for the Latin text only');
  assert.deepEqual((await shown('zh-Hant')).blank, ['新細明體', 'Chinese (Trad.)'], 'zh-HK has a name');
  const en = await shown('en');
  assert.deepEqual([en.blank, en.latin, en.ea], [['Times New Roman', 'English (U.S.)'], ['Times New Roman', 'English (U.S.)'], ['Times New Roman', 'English (U.S.)']],
    'English: no East Asian font or language to show');
});

test('Ledger, Japanese and Korean: their Excel\'s default font on A4', async () => {
  const ja = await ledger('ja'), ko = await ledger('ko');
  assert.match(ja['xl/styles.xml'], /<fonts[^>]*>\s*<font>.*?<name val="ＭＳ Ｐゴシック"\/>.*?<charset val="128"\/>/s);
  assert.match(ja['xl/styles.xml'], /<fonts[^>]*>\s*<font>\s*<sz val="11"\/>/s);
  assert.match(ko['xl/styles.xml'], /<fonts[^>]*>\s*<font>.*?<name val="맑은 고딕"\/>.*?<charset val="129"\/>/s);
  assert.match(ja['xl/worksheets/sheet1.xml'], /<pageSetup [^>]*paperSize="9"/);
  assert.match(ja['xl/worksheets/sheet1.xml'], /<sheetFormatPr [^>]*defaultRowHeight="13.5"/);
});

test('Lectern, Japanese: text in Japanese, the theme\'s Japanese font', async () => {
  const x = await lectern('ja');
  assert.match(x['ppt/presentation.xml'], /<p:defaultTextStyle><a:defPPr><a:defRPr lang="ja-JP"\/>/);
  assert.match(x['ppt/theme/theme1.xml'], /<a:minorFont>.*<a:font script="Jpan" typeface="ＭＳ Ｐゴシック"\/>.*<\/a:minorFont>/s);
  const en = await lectern('en');
  assert.match(en['ppt/presentation.xml'], /<a:defRPr lang="en-US"\/>/);
  assert.doesNotMatch(en['ppt/theme/theme1.xml'], /script="Jpan"/);
});

test('Lectern design changes keep the target slide’s East Asian fonts through a PPTX save', async () => {
  const page = await openPage('lectern', { lang: 'ja', persistence: false, automated: true });
  try {
    const inspect = async () => {
      const M = L.model;
      const fonts = (p, i = 0) => {
        const f = M.design(p, p.slides[i]).fonts;
        return [f.majorEA, f.minorEA];
      };
      const write = async p => new Uint8Array(await (await L.pptx.write(p)).arrayBuffer());
      const pres = M.newPresentation();
      M.applyDesign(pres, 'meadow');
      const bytes = await write(pres), pkg = await L.opc.open(bytes);
      const theme = pkg.xml('ppt/theme/theme1.xml');
      const japanese = ['majorFont', 'minorFont'].map(role => {
        const group = theme.getElementsByTagNameNS('*', role)[0];
        return [...group.getElementsByTagNameNS('*', 'font')].find(f => f.getAttribute('script') === 'Jpan')?.getAttribute('typeface');
      });
      const own = M.buildDesign('night', pres.W, pres.H);
      Object.assign(own.fonts, { majorEA: 'Yu Gothic', minorEA: 'Yu Mincho' });
      pres.designs[own.id] = own;
      pres.slides.push(M.newSlide(pres, 'text', own.id));
      const opened = await L.pptx.read(await write(pres));
      M.applyDesign(opened, 'ocean', [opened.slides[1]]);
      const again = await L.pptx.read(await write(opened));
      const english = await L.pptx.read(await write(M.newPresentation({ locale: false })));
      M.applyDesign(english, 'azure');
      const fallback = await L.pptx.read(await write(english));
      const others = [M.newPresentation({ design: 'night' }), ...L.templates.list().map(t => L.templates.build(t.id))];
      await L.app.openFile(new File(['Outline\n\tTopic'], 'Outline.txt'));
      others.push(L.pres);
      return { design: fonts(pres), japanese, first: fonts(again), second: fonts(again, 1),
        fallback: fonts(fallback), otherPaths: others.map(p => [p.lang, ...fonts(p)]) };
    };
    const result = await page.evaluate('(' + inspect + ')()');
    const ja = ['ＭＳ Ｐゴシック', 'ＭＳ Ｐゴシック'];
    assert.deepEqual(result.design, ja);
    assert.deepEqual(result.japanese, ja, 'both roles in the saved theme');
    assert.deepEqual(result.first, ja, 'unselected slide');
    assert.deepEqual(result.second, ['Yu Gothic', 'Yu Mincho'], 'the first selected slide supplied both roles');
    assert.deepEqual(result.fallback, ja, 'a presentation without East Asian fonts takes the UI default');
    for (const path of result.otherPaths) assert.deepEqual(path, ['ja-JP', ...ja], 'design list, Start Center templates and text outline');
  } finally { await page.close(); }
});

async function quireTemplates(lang) {
  const page = await openPage('quire', { lang, persistence: false, automated: true });
  try {
    const inspect = async () => {
      const { D } = L, loc = L.locale.get();
      const saved = async doc => {
        const pkg = await L.opc.open(await (await L.docx.write(doc)).arrayBuffer());
        const node = (parent, name) => parent.getElementsByTagNameNS('*', name)[0];
        const docXml = pkg.xml('word/document.xml'), settings = pkg.xml('word/settings.xml');
        const styles = pkg.xml('word/styles.xml'), theme = pkg.xml('word/theme/theme1.xml');
        const part = (parent, name) => { const el = node(parent, name); return el ? L.opc.raw(el) : ''; };
        return { paper: part(docXml, 'pgSz'), grid: part(docXml, 'docGrid'), defaults: part(styles, 'docDefaults'),
          themeLang: part(settings, 'themeFontLang'), major: part(theme, 'majorFont'), minor: part(theme, 'minorFont') };
      };
      const blank = await saved(D.doc), margins = d => ['mt', 'mr', 'mb', 'ml'].map(k => d.sect[k]);
      const blankMargins = margins(D.doc);
      const results = [];
      for (const { id } of L.templates.list()) {
        L.templates.create(id); // File > New, using the same builder as the Start Center
        const d = D.doc, custom = {}, wrongHeadingFonts = [];
        for (const [key, s] of Object.entries(d.styles)) {
          if (!s.custom) continue;
          const { fontEA, ...latin } = s.rPr;
          custom[key] = { rPr: latin, pPr: s.pPr };
          if (loc.themeEA && /^(Arial|Tahoma)$/.test(latin.font) && (fontEA || d.defaults.rPr.fontEA) !== loc.themeEA.major) wrongHeadingFonts.push(key);
        }
        results.push({ id, saved: await saved(d), margins: margins(d), custom, wrongHeadingFonts,
          blocks: d.main.blocks.length, normal: d.styles.Normal });
      }
      return { blank, blankMargins, results };
    };
    return await page.evaluate('(' + inspect + ')()');
  } finally { await page.close(); }
}

test('Quire templates save the language’s paper, grid, fonts and theme while keeping each template’s design', async () => {
  const en = await quireTemplates('en');
  for (const lang of ['ja', 'ko', 'zh-CN', 'zh-Hant', 'fr', 'de']) {
    const { blank, blankMargins, results } = await quireTemplates(lang);
    assert.match(blank.paper, /pgSz/);
    assert.match(blank.defaults, /rFonts/);
    assert.match(blank.themeLang, /themeFontLang/);
    for (const entry of results) {
      const message = lang + ' ' + entry.id, reference = en.results.find(t => t.id === entry.id);
      assert.deepEqual(entry.saved, blank, message + ' OOXML defaults');
      assert.deepEqual(entry.custom, reference.custom, message + ' template styles');
      assert.deepEqual(entry.wrongHeadingFonts, [], message + ' heading fonts');
      assert.equal(entry.blocks, reference.blocks, message + ' content');
      const margins = [...blankMargins];
      if (entry.id === 'letter') margins[0] = 72;
      if (entry.id === 'letter' || entry.id === 'resume') margins[1] = margins[3] = 90;
      if (entry.id === 'flyer') margins[0] = margins[2] = 54;
      assert.deepEqual(entry.margins, margins, message + ' margins');
      if (entry.id === 'report') {
        assert.equal(entry.normal.rPr.font, 'Garamond');
        assert.equal(entry.normal.rPr.sz, 11);
        assert.deepEqual(entry.normal.pPr.sp, reference.normal.pPr.sp);
        if (lang === 'ja') assert.equal(entry.normal.pPr.jc, 'both', 'Report keeps Japanese paragraph defaults');
      }
    }
  }
});
