/* Browser-only binary Word -> WordprocessingML package orchestration. */
(function (root) {
  'use strict';
  const B = root.L.legacy;
  const D = B.doc;
  const N = s => s;
  D.convert = async function (bytes, options = {}) {
    const cfb = D.decrypt(B.cfb(bytes), options);
    const raw = cfb.stream('WordDocument');
    if (!raw) B.fail('DOC_STREAM', N('The WordDocument stream is missing'));
    const fib = D.fib(raw);
    if (fib.version !== 'Word97+')
      B.fail('DOC_VERSION', N('Word 6, Word 95 and earlier documents cannot be opened yet'));
    const book = D.structure(D.read(cfb), cfb, []);
    const losses = book.losses;
    const pkg = new B.Package();
    const main = 'word/document.xml';
    const ct = 'application/vnd.openxmlformats-officedocument.wordprocessingml.';
    const macros = B.vba.read(cfb, 'Macros', losses, 'Document');
    D.styles(book);
    D.lists(book);
    D.settings(book);
    D.fields.prepare(book);
    D.markup.prepare(book);
    book.textStats = { paragraphs: 0, runs: 0, characters: 0, deletedCharacters: 0,
      instructionCharacters: 0, objects: 0 };
    await D.draw.prepare(book, pkg);
    const body = D.writeStories(book, pkg);
    const template = fib.template || String(options.extension || '').replace(/^\./, '').toLowerCase() === 'dot';
    const type = (template ? 'dot' : 'doc') + (macros ? 'm' : 'x');
    const contentType = macros
      ? 'application/vnd.ms-word.' + (template ? 'template' : 'document') + '.macroEnabled.main+xml'
      : ct + (template ? 'template' : 'document') + '.main+xml';
    pkg.add(main, D.part('document', B.el('w:body', {}, body)), contentType);
    pkg.rel('', 'officeDocument', main);
    const parts = [['styles', D.stylesXML(book)], ['settings', D.settingsXML(book)],
      ['fontTable', D.fontsXML(book)], ...(book.lists.length ? [['numbering', D.numberingXML(book)]] : [])];
    for (const [name, data] of parts) {
      pkg.add('word/' + name + '.xml', data, ct + name + '+xml');
      pkg.rel(main, name, name + '.xml');
    }
    losses.push(...B.addProperties(pkg, cfb));
    D.markup.writeComments(book, pkg);
    D.markup.finish(book);
    D.draw.finish(book);
    if (book.settings.protected)
      D.notice(book, 'legacy-word-protection', N('Document editing restrictions were not kept'), 'Document', true);
    if (book.settings.embeddedFonts)
      D.notice(book, 'legacy-embedded-fonts', N('Embedded fonts were not kept'), 'Document', true);
    if (B.vba.add(pkg, main, macros)) B.vba.wordData(book, pkg);
    return { pkg: await pkg.finish(), type, losses, stats: {
      ...book.textStats, ...book.fieldStats, ...book.markupStats, ...book.tableStats, ...book.drawStats,
      sections: book.sections.length, styles: book.styles.filter(Boolean).length, fonts: book.fonts.length,
      lists: book.lists.length, footnotes: book.footnotes.length, endnotes: book.endnotes.length,
      sourceTextCharacters: book.stories.main, storedPages: book.settings.storedPages,
      ...(macros ? { macros: { streams: macros.streams, bytes: macros.bytes.length } } : {})
    } };
  };
})(typeof window !== 'undefined' ? window : globalThis);
