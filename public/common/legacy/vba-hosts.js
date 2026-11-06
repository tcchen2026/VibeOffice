/* Binary Office host bindings for preserved VBA projects.
 * Derived from Apache OpenOffice main/sc/source/filter/excel/xiroot.cxx (Apache-2.0).
 * Derived from Apache POI hslf/record/{VBAInfoAtom,ExOleObjStg}.java and
 * hwpf/model/types/DOPAbstractType.java (Apache-2.0).
 * MS-XLS CodeName; MS-PPT 2.4.10–11, 2.10.40–42; MS-DOC Dop97, Tcg and Mcd;
 * MS-OFFMACRO2 vbaSuppData. No VBA source is translated or executed. */
(function (root) {
  'use strict';
  const B = root.L.legacy;
  const V = B.vba;
  const N = s => s;
  const CODE_NAME = 0x01ba;
  const DOC_INFO_LIST = 2000;
  const VBA_INFO = 1023;
  const VBA_INFO_ATOM = 1024;
  const VBA_STORAGE = 4113;
  const WORD_COMMANDS = 24;
  const WORD_DOP = 31;
  const WORD_EVENTS_OFFSET = 0x1b2;
  const WORD_VARS = 60;
  const VBA_SIGNATURE = 0x0018;
  const DOCUMENT_SUMMARY = 'd5cdd502-2e9c-101b-9397-08002b2cf9ae';
  const COMMAND_MACROS = 0x01;
  const COMMAND_ALLOCATED = 0x02;
  const COMMAND_KEYS = 0x03;
  const COMMAND_INVALID_KEYS = 0x04;
  const COMMAND_STRINGS = 0x10;
  const COMMAND_NAMES = 0x11;
  const COMMAND_TOOLBARS = 0x12;
  const COMMAND_END = 0x40;
  const WNE = 'http://schemas.microsoft.com/office/word/2006/wordml';
  const MICROSOFT_REL = 'http://schemas.microsoft.com/office/2006/relationships/';
  const bad = detail => B.fail('VBA_HOST', 'Invalid macro host information', detail);

  V.signatureNotice = function (cfb, losses, where) {
    // MS-OSHARED 2.3.3.2.2.1: this property holds the VBA signature in XLS/PPT.
    // Conversion preserves the project but does not transfer its signature wrapper.
    try {
      const sets = B.properties(cfb.stream('\x05DocumentSummaryInformation'));
      const signature = sets.find(set => set.fmtid === DOCUMENT_SUMMARY)?.values[VBA_SIGNATURE];
      if (signature?.value?.length)
        V.notice(losses, N('The macro digital signature was not kept'), where);
    } catch (error) {
      if (!['LegacyError', 'CFBError'].includes(error.name)) throw error;
      // Optional metadata damage is reported by the property reader, not as lost VBA.
    }
  };

  V.excel = function (book, cfb, losses) {
    const project = V.read(cfb, '_VBA_PROJECT_CUR', losses, 'Workbook');
    if (!project) return null;
    V.signatureNotice(cfb, losses, 'Workbook');
    const codeName = (records, where) => {
      const record = records.find(r => r.id === CODE_NAME);
      if (!record) return null;
      try {
        const value = B.xls.string(record.r, 0, book.version === 'BIFF8', book.codepage).value;
        if (value.length > 31 || /[\x00-\x1f\ufffe\uffff]/.test(value)) bad('Invalid code name');
        return value || null;
      } catch (error) {
        if (error.name !== 'LegacyError') throw error;
        V.notice(losses, N('A macro connection to this workbook or sheet was not kept'), where, error.code);
        return null;
      }
    };
    project.workbookName = codeName(book.globals, 'Workbook');
    project.sheetNames = book.sheets.map(sheet => codeName(sheet.records, sheet.name));
    return project;
  };

  V.powerpoint = async function (book) {
    const P = B.ppt;
    const where = 'Presentation';
    try {
      const records = book.documentRecords.flatMap(r => r.type === DOC_INFO_LIST ? P.children(r, () => {}) : [r]);
      const info = records.find(r => r.type === VBA_INFO);
      if (!info) return null;
      const atom = P.children(info).find(r => r.type === VBA_INFO_ATOM)?.r;
      if (!atom) bad('Missing VBA information atom');
      if (!atom.u32(4)) return null;
      const storage = book.get(atom.u32(0));
      if (storage.type !== VBA_STORAGE || ![0, 1].includes(storage.instance)) bad('Invalid VBA persist object');
      const bytes = storage.instance
        ? await B.images.inflate(storage.r.sub(4, storage.r.length - 4), storage.r.u32(0), B.limits.bytes)
        : storage.r.bytes;
      const project = V.read(B.cfb(bytes), '', book.losses, where);
      // PowerPoint already stores a standalone project CFB. Preserve that whole
      // storage, including its allocation and caches, after checking its streams.
      if (project) project.bytes = bytes;
      return project;
    } catch (error) {
      return V.damaged(book.losses, where, error);
    }
  };

  V.wordCommands = function (book) {
    const D = B.doc;
    const r = D.table(book, WORD_COMMANDS, 'Macro commands');
    if (!r.length) return '';
    if (r.u8(0) !== 255) bad('Unknown command-table version');
    const macros = [];
    const strings = [];
    const names = new Map();
    let p = 1;
    while (p < r.length) {
      const kind = r.u8(p++);
      if (kind === COMMAND_END) break;
      if (kind === COMMAND_MACROS) {
        const count = r.u32(p);
        p += 4;
        r.check(p, count * 24);
        for (let i = 0; i < count; i++, p += 24)
          macros.push({ cmg: r.u8(p), encrypt: r.u8(p + 1), id: r.u16(p + 2), name: r.u16(p + 4) });
      } else if (kind === COMMAND_STRINGS) {
        const count = r.u16(p + 2);
        const extra = r.u16(p + 4);
        if (r.u16(p) !== 0xffff) bad('Unknown command string encoding');
        p += 6;
        for (let i = 0; i < count; i++) {
          const length = r.u16(p);
          p += 2;
          if (length === 0xffff) strings.push('');
          else {
            strings.push(D.unicode(r, p, length * 2));
            r.check(p, length * 2 + extra);
            p += length * 2 + extra;
          }
        }
      } else if (kind === COMMAND_NAMES) {
        const count = r.u16(p);
        p += 2;
        for (let i = 0; i < count; i++) {
          const id = r.u16(p);
          const length = r.u16(p + 2);
          p += 4;
          const name = D.unicode(r, p, length * 2);
          r.check(p, length * 2 + 2);
          p += length * 2 + 2;
          names.set(id, name);
        }
      } else if ([COMMAND_ALLOCATED, COMMAND_KEYS, COMMAND_INVALID_KEYS].includes(kind)) {
        const count = r.u32(p);
        p += 4;
        const size = count * (kind === COMMAND_ALLOCATED ? 4 : 14);
        r.check(p, size);
        p += size;
        if (count && kind === COMMAND_KEYS)
          V.notice(book.losses, N('Custom keyboard shortcuts were not kept'), 'Document');
      } else if (kind === COMMAND_TOOLBARS) {
        V.notice(book.losses, N('Custom toolbars were not kept'), 'Document');
        break;
      } else bad('Unknown command-table entry');
    }
    const hex = n => n.toString(16).padStart(2, '0').toUpperCase();
    return macros.map(macro => {
      const name = strings[macro.name];
      const macroName = names.get(macro.id);
      if (name == null || macroName == null) bad('Missing macro command name');
      return B.el('wne:mcd', { 'wne:macroName': macroName, 'wne:name': name,
        'wne:bEncrypt': hex(macro.encrypt), 'wne:cmg': hex(macro.cmg) });
    }).join('');
  };

  V.wordData = function (book, pkg) {
    const D = B.doc;
    let commands = '';
    let events = '';
    try {
      commands = V.wordCommands(book);
    } catch (error) {
      if (error.name !== 'LegacyError') throw error;
      V.notice(book.losses, N('Some macro command names were not kept'), 'Document', error.code);
    }
    try {
      const dop = D.table(book, WORD_DOP, 'Macro document events');
      const flags = dop.length >= WORD_EVENTS_OFFSET + 4 ? dop.u32(WORD_EVENTS_OFFSET) : 0;
      const names = ['New', 'Open', 'Close', 'Sync', 'XmlAfterInsert', 'XmlBeforeDelete', null, null,
        'ContentControlAfterInsert', 'ContentControlBeforeDelete', 'ContentControlOnExit', 'ContentControlOnEnter',
        'StoreUpdate', 'ContentControlContentUpdate', 'BuildingBlockAfterInsert'];
      events = names.map((name, bit) => name && flags & (1 << bit) ? B.el('wne:eventDoc' + name) : '').join('');
    } catch (error) {
      if (error.name !== 'LegacyError') throw error;
      V.notice(book.losses, N('Macro document events could not be read'), 'Document', error.code);
    }
    pkg.add('word/vbaData.xml', B.xmlDeclaration + B.el('wne:vbaSuppData', { 'xmlns:wne': WNE },
      (events ? B.el('wne:docEvents', {}, events) : '') + (commands ? B.el('wne:mcds', {}, commands) : '')),
    'application/vnd.ms-word.vbaData+xml');
    pkg.rel('word/vbaProject.bin', MICROSOFT_REL + 'wordVbaData', 'vbaData.xml');
    // Word stores VBA signatures in binary document variables. Their signed
    // host data is not carried into this conversion; never publish a stale claim.
    try {
      if (!book.pair(WORD_VARS).length) return;
      const table = D.table(book, WORD_VARS, 'Macro signatures');
      const state = {};
      const variables = D.markup.strings(table, state);
      let p = state.end;
      for (const entry of variables) {
        const length = table.u16(p);
        p += 2;
        table.check(p, length * 2);
        if (length && /^(sign|sigagile|sigv3)$/i.test(entry.text)) {
          V.notice(book.losses, N('The macro digital signature was not kept'), 'Document');
          break;
        }
        p += length * 2;
      }
    } catch (error) {
      if (error.name !== 'LegacyError') throw error;
    }
  };
})(typeof window !== 'undefined' ? window : globalThis);
