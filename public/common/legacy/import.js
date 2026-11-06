/* Office 97-2003 files (.doc, .xls, .ppt and their templates): the converters in this folder load when such a file
 * is opened, and turn it into the OOXML package Office would write for it, which the app then reads as usual (it is
 * saved as .docx, .xlsx or .pptx; nothing is written in the old formats). Loaded by every app after common/crypto.js;
 * the converters need zip, crypto, xml and opc. */
(function (root) {
  'use strict';
  const L = root.L;
  /* the converters in load order (tools/legacy/load.mjs reads this list too) */
  const FILES = ['base', 'codepages', 'properties', 'xls-read', 'doc-read', 'ppt-read', 'inspect',
    'package', 'property-parts', 'images', 'officeart', 'native-geometry', 'shape-path', 'shapes', 'xml-check',
    'xls-style-ext', 'xls-style', 'xls-sheet', 'xls-functions', 'xls-formula', 'xls-formula-book',
    'xls-misc', 'xls-dxf', 'xls-rules', 'xls-comments', 'xls-write',
    'doc-sprm', 'doc-structure', 'doc-style', 'doc-lists', 'doc-xml', 'doc-settings', 'doc-text', 'doc-write',
    'ppt-structure', 'ppt-package', 'ppt-theme', 'ppt-text-read', 'ppt-text', 'ppt-text-design', 'ppt-text-write',
    'ppt-shapes', 'ppt-tables', 'ppt-links', 'ppt-features', 'ppt-write'];
  const N = (s) => s;
  const KIND = { doc: N('Word 97-2003 document'), xls: N('Excel 97-2003 workbook'), ppt: N('PowerPoint 97-2003 presentation') };
  const OTHER = { doc: N('This is a Word 97-2003 document, which {0} does not open.'), xls: N('This is an Excel 97-2003 workbook, which {0} does not open.'), ppt: N('This is a PowerPoint 97-2003 presentation, which {0} does not open.') };
  const here = typeof document !== 'undefined' && document.currentScript ? new URL('.', document.currentScript.src).href : '';
  let loading = null;
  /** the converters, loaded once, one after another (each attaches to L.legacy) */
  const load = () => loading || (loading = FILES.reduce((p, f) => p.then(() => new Promise((resolve, reject) => {
    const s = document.createElement('script');
    s.src = here + f + '.js';
    s.onload = resolve;
    s.onerror = () => { loading = null; reject(new Error(L.t('The converter for Office 97-2003 files could not be loaded.'))); };
    document.head.appendChild(s);
  })), Promise.resolve()));

  L.legacyImport = {
    FILES,
    load,
    kindName: (format) => L.t(KIND[format] || format),
    /** bytes of an OLE compound file -> { pkg, type, losses, format }; want is the app's format ('doc', 'xls',
        'ppt'). A refusal throws with code 'legacy' and a message for the user. */
    async convert(bytes, name, want) {
      await load();
      const B = L.legacy, refuse = (message) => { const e = new Error(message); e.code = 'legacy'; return e; };
      const u8 = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
      let sniff;
      try { sniff = B.sniff(u8); } catch (e) { throw refuse(L.t('The file is damaged or is not an Office 97-2003 file.')); }
      if (sniff.format !== want) {
        throw refuse(OTHER[sniff.format] ? L.t(OTHER[sniff.format], L.APP) : L.t('The file is damaged or is not an Office 97-2003 file.'));
      }
      const extension = (/\.([a-z]+)$/i.exec(name || '') || [])[1];
      try {
        const r = await B[want].convert(u8, { extension });
        return Object.assign({ format: want }, r);
      } catch (e) {
        if (!e || !['LegacyError', 'CFBError'].includes(e.name)) throw e;
        if (e.code === 'PASSWORD_REQUIRED' || e.code === 'ENCRYPTION_UNSUPPORTED') throw refuse(L.t('This {0} is protected with a password. {1} cannot open password-protected Office 97-2003 files yet.', L.legacyImport.kindName(want), L.APP));
        throw refuse(L.t(e.message));
      }
    },
    /** what the user is told after a conversion, or '' when nothing they had was left out */
    notice(result, fileName) {
      const what = [...new Set((result.losses || []).filter((l) => l.notify === true).map((l) => L.t(l.what)))];
      if (!what.length) return '';
      return L.t('{0} was converted from the {1} format. Not converted:', fileName, L.legacyImport.kindName(result.format)) + '\n' + what.map((w) => '• ' + w).join('\n');
    },
  };
})(typeof window !== 'undefined' ? window : globalThis);
