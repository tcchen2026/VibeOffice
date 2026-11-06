/* Binary Office import foundations. Byte bounds and XML rules are independent implementations.
 * Formats: Microsoft Open Specifications MS-CFB, MS-OLEPS and ECMA-376. No DOM is needed. */
(function (root) {
  'use strict';
  const N = s => s; // UI refusal text, translated where it is shown.
  const L = root.L || (root.L = {}),
    B = L.legacy || (L.legacy = {});
  // Windows LCIDs stored in binary Office character properties.
  B.languageIds = Object.freeze({
    1025: 'ar-SA',
    1028: 'zh-TW',
    1029: 'cs-CZ',
    1030: 'da-DK',
    1031: 'de-DE',
    1032: 'el-GR',
    1033: 'en-US',
    1034: 'es-ES',
    1035: 'fi-FI',
    1036: 'fr-FR',
    1037: 'he-IL',
    1038: 'hu-HU',
    1040: 'it-IT',
    1041: 'ja-JP',
    1042: 'ko-KR',
    1043: 'nl-NL',
    1044: 'nb-NO',
    1045: 'pl-PL',
    1046: 'pt-BR',
    1048: 'ro-RO',
    1049: 'ru-RU',
    1050: 'hr-HR',
    1051: 'sk-SK',
    1053: 'sv-SE',
    1054: 'th-TH',
    1055: 'tr-TR',
    1057: 'id-ID',
    1058: 'uk-UA',
    1060: 'sl-SI',
    1061: 'et-EE',
    1062: 'lv-LV',
    1063: 'lt-LT',
    1066: 'vi-VN',
    1081: 'hi-IN',
    2052: 'zh-CN',
    2057: 'en-GB',
    2070: 'pt-PT',
    3076: 'zh-HK',
    3082: 'es-ES',
    4105: 'en-CA',
    5129: 'en-NZ'
  });
  B.limits = Object.freeze({
    bytes: 256 * 1024 * 1024,
    text: 32 * 1024 * 1024,
    records: 2000000,
    depth: 128,
    cells: 2000000
  });
  B.fail = function (code, message, where) {
    const e = new Error(message);
    e.name = 'LegacyError';
    e.code = code;
    if (where != null) e.where = where;
    throw e;
  };
  B.Reader = class {
    constructor(bytes, label = 'stream') {
      if (!(bytes instanceof Uint8Array)) B.fail('INVALID_BYTES', 'Expected a byte array', label);
      if (bytes.length > B.limits.bytes) B.fail('SIZE_LIMIT', N('The file is too large to import'), label);
      this.source = bytes;
      this.start = 0;
      this.length = bytes.length;
      this.label = label;
      this.backing = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    }
    get view() {
      if (!this.localView) {
        this.localView =
          this.start === 0 && this.length === this.source.length
            ? this.backing
            : new DataView(this.source.buffer, this.source.byteOffset + this.start, this.length);
      }
      return this.localView;
    }
    get bytes() {
      return this.start === 0 && this.length === this.source.length ? this.source : this.sub(0, this.length);
    }
    check(p, n) {
      if (!Number.isSafeInteger(p) || !Number.isSafeInteger(n) || p < 0 || n < 0 || p > this.length - n)
        B.fail(
          'TRUNCATED',
          N('The file contains an invalid offset or length'),
          this.label + ':' + (this.start + p)
        );
      return p;
    }
    u8(p) {
      return this.backing.getUint8(this.start + this.check(p, 1));
    }
    u16(p) {
      return this.backing.getUint16(this.start + this.check(p, 2), true);
    }
    i16(p) {
      return this.backing.getInt16(this.start + this.check(p, 2), true);
    }
    u32(p) {
      return this.backing.getUint32(this.start + this.check(p, 4), true);
    }
    i32(p) {
      return this.backing.getInt32(this.start + this.check(p, 4), true);
    }
    f32(p) {
      return this.backing.getFloat32(this.start + this.check(p, 4), true);
    }
    f64(p) {
      return this.backing.getFloat64(this.start + this.check(p, 8), true);
    }
    sub(p, n) {
      this.check(p, n);
      return this.source.subarray(this.start + p, this.start + p + n);
    }
    reader(p, n, label = this.label) {
      this.check(p, n);
      // Subreaders retain bounds but share the backing view. Large BIFF streams
      // have hundreds of thousands of records; each needs no extra byte/view objects.
      const child = Object.create(B.Reader.prototype);
      child.source = this.source;
      child.start = this.start + p;
      child.length = n;
      child.label = label;
      child.backing = this.backing;
      return child;
    }
    text(p, n, cp = 1252) {
      return B.decode(this.sub(p, n), cp);
    }
  };
  const decoders = new Map();
  const labels = {
    1200: 'utf-16le',
    1201: 'utf-16be',
    65001: 'utf-8',
    932: 'shift_jis',
    936: 'gbk',
    949: 'euc-kr',
    950: 'big5',
    874: 'windows-874',
    10000: 'macintosh',
    10007: 'x-mac-cyrillic',
    28591: 'windows-1252'
  };
  B.decode = function (bytes, codepage = 1252) {
    if (B.codepages?.[codepage]) {
      const table = B.codepages[codepage];
      let out = '';
      for (let p = 0; p < bytes.length; p += 8192) {
        const chunk = bytes.subarray(p, p + 8192);
        out += Array.from(chunk, b => table[b]).join('');
      }
      return out;
    }
    const label = labels[codepage] || (codepage >= 1250 && codepage <= 1258 ? 'windows-' + codepage : null);
    if (!label) B.fail('CODEPAGE_UNSUPPORTED', N('This file uses an unsupported text encoding'), codepage);
    if (!decoders.has(label)) {
      try {
        decoders.set(label, new TextDecoder(label));
      } catch {
        B.fail('CODEPAGE_UNSUPPORTED', N('This browser cannot decode the file text encoding'), label);
      }
    }
    return decoders.get(label).decode(bytes);
  };
  // BIFF8 compressed Unicode stores low UTF-16 bytes, not Windows-1252.
  B.latin16 = function (bytes) {
    let s = '';
    for (let p = 0; p < bytes.length; p += 8192) s += String.fromCharCode(...bytes.subarray(p, p + 8192));
    return s;
  };
  B.xml = function (value) {
    return String(value)
      .replace(
        /[\u0000-\u0008\u000b\u000c\u000e-\u001f\ufffe\uffff]|[\ud800-\udbff](?![\udc00-\udfff])|(?<![\ud800-\udbff])[\udc00-\udfff]/g,
        '\ufffd'
      )
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  };
  // SpreadsheetML/ST_Xstring escapes controls and literal escape sequences (MS-XLSX §2.5.2).
  B.xstring = value =>
    B.xml(
      String(value)
        .replace(/_x[0-9a-fA-F]{4}_/g, s => '_x005F_' + s.slice(1))
        .replace(
          /[\x00-\x08\x0b\x0c\x0e-\x1f\r]/g,
          c => '_x' + c.charCodeAt(0).toString(16).padStart(4, '0').toUpperCase() + '_'
        )
    );
  B.cfb = bytes => {
    if (!L.officeCrypto?.cfb) B.fail('DEPENDENCY', 'Load common/crypto.js before the legacy importers');
    return L.officeCrypto.cfb(bytes, { maxBytes: B.limits.bytes });
  };
  // Notification is a host/converter decision: opt in only for lost content or functionality.
  // Property-set and metadata problems are always silent diagnostics.
  B.loss = (id, what, where, action = 'drop', place, notify = false) => ({
    id,
    what,
    where,
    action,
    notify: notify === true,
    ...(place ? { place } : {})
  });
})(typeof window !== 'undefined' ? window : globalThis);
