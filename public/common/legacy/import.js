/* Office 97-2003 files (.doc, .xls, .ppt and their templates): the converters in this folder load when such a file
 * is opened, and turn it into the OOXML package Office would write for it, which the app then reads as usual (it is
 * saved as .docx, .xlsx or .pptx; nothing is written in the old formats). Loaded by every app after common/crypto.js;
 * the converters need zip, crypto, xml and opc. */
(function (root) {
  'use strict';
  const L = root.L;
  /* the converters in load order (tools/legacy/load.mjs reads this list too) */
  const FILES = ['base', 'codepages', 'properties', 'encryption-crypto', 'encryption',
    'xls-read', 'doc-read', 'ppt-read', 'xls-crypto', 'doc-crypto', 'ppt-crypto', 'inspect',
    'package', 'property-parts', 'images', 'officeart', 'native-geometry', 'shape-path', 'shapes', 'xml-check',
    'xls-style-ext', 'xls-style', 'xls-sheet', 'xls-functions', 'xls-formula', 'xls-formula-book',
    'xls-misc', 'xls-dxf', 'xls-rules', 'xls-draw-read', 'xls-comments', 'xls-draw-write',
    'xls-chart-read', 'xls-chart-style', 'xls-chart-write', 'xls-write',
    'doc-sprm', 'doc-structure', 'doc-style', 'doc-lists', 'doc-xml', 'doc-settings',
    'doc-fields', 'doc-markup', 'doc-table-props', 'doc-tables', 'doc-draw-read', 'doc-draw-write', 'doc-text', 'doc-write',
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
        'ppt'), opts.password the password to open, when the user gave one. A file with a password to open throws
        with code 'password' (none given; Office's default password is tried first, as Office opens such files
        without asking) or 'badpassword', as the apps' readers do for encrypted .docx/.xlsx/.pptx; a refusal throws
        with code 'legacy' and a message for the user. The package returned is not encrypted. */
    async convert(bytes, name, want, opts) {
      await load();
      const B = L.legacy, refuse = (message, code) => { const e = new Error(message); e.code = code || 'legacy'; return e; };
      const u8 = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
      let sniff;
      try { sniff = B.sniff(u8); } catch (e) { throw refuse(L.t('The file is damaged or is not an Office 97-2003 file.')); }
      if (sniff.format !== want) {
        throw refuse(OTHER[sniff.format] ? L.t(OTHER[sniff.format], L.APP) : L.t('The file is damaged or is not an Office 97-2003 file.'));
      }
      const extension = (/\.([a-z]+)$/i.exec(name || '') || [])[1], password = opts && opts.password;
      try {
        const r = await B[want].convert(u8, password == null ? { extension } : { extension, password });
        return Object.assign({ format: want }, r);
      } catch (e) {
        if (!e || !['LegacyError', 'CFBError'].includes(e.name)) throw e;
        if (e.code === 'PASSWORD_REQUIRED') throw refuse(L.N('This file is password-protected.'), 'password');
        if (e.code === 'WRONG_PASSWORD') throw refuse(L.N('The password you supplied is not correct. Verify that the CAPS LOCK key is off and be sure to use the correct capitalization.'), 'badpassword');
        if (e.code === 'ENCRYPTION_UNSUPPORTED') throw refuse(L.t('This {0} is protected with a kind of password {1} cannot open.', L.legacyImport.kindName(want), L.APP));
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
