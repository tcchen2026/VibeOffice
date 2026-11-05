/* OLE property sets: MS-OLEPS §§2.5–2.21; Office's unaligned vectors: MS-OSHARED §2.3.3.
 * Derived from Apache POI hpsf/TypedPropertyValue.java, Vector.java,
 * CodePageString.java and UnicodeString.java (Apache-2.0): typed values and tolerant padding.
 * Each property is bounded by the next offset; a damaged optional value cannot hide its siblings. */
(function (root) {
  'use strict';
  const N = s => s; // UI refusal text, translated where it is shown.
  const B = root.L.legacy;
  const guid = (r, p) => {
    const hex = (v, n) => v.toString(16).padStart(n, '0');
    return hex(r.u32(p), 8) + '-' + hex(r.u16(p + 4), 4) + '-' + hex(r.u16(p + 6), 4) + '-' +
      [...r.sub(p + 8, 2)].map(v => hex(v, 2)).join('') + '-' + [...r.sub(p + 10, 6)].map(v => hex(v, 2)).join('');
  };
  const date = ms => Number.isFinite(ms) && Math.abs(ms) <= 8640000000000000 ? new Date(ms).toISOString() : null;
  B.properties = function (bytes, options = {}) {
    if (!bytes) return [];
    const r = new B.Reader(bytes, 'OLE properties');
    if (r.u16(0) !== 0xfffe) B.fail('PROPERTIES_ENDIAN', N('Unsupported property-set byte order'));
    const count = r.u32(24); if (count > 16) B.fail('PROPERTIES_LIMIT', N('Too many property sets'));
    r.check(28, count * 20);
    const sets = [];
    for (let i = 0; i < count; i++) {
      const fmtid = guid(r, 28 + i * 20), values = {}, dictionary = {}, issues = [];
      const set = { fmtid, codepage: 1252, values, dictionary, issues }; sets.push(set);
      try {
        const base = r.u32(44 + i * 20), size = r.u32(base), s = r.reader(base, size), n = s.u32(4);
        if (n > 100000) B.fail('PROPERTIES_LIMIT', N('Too many properties')); s.check(8, n * 8);
        const offsets = new Map(), boundaries = new Set([size]);
        for (let j = 0; j < n; j++) {
          const id = s.u32(8 + j * 8), at = s.u32(12 + j * 8);
          if (at < 8 + n * 8 || at > size - 4 || offsets.has(id)) { issues.push({ id, code: 'PROPERTIES_OFFSET' }); continue; }
          offsets.set(id, at); boundaries.add(at);
        }
        const ends = new Map(), sorted = [...boundaries].sort((a, b) => a - b);
        for (let j = 0; j < sorted.length - 1; j++) ends.set(sorted[j], sorted[j + 1]);
        const cpAt = offsets.get(1);
        if (cpAt != null && ends.get(cpAt) - cpAt >= 6) set.codepage = s.u16(cpAt + 4) || 1252;
        const officeSummary = options.kind === 'documentSummary' && i === 0 || fmtid === 'd5cdd502-2e9c-101b-9397-08002b2cf9ae';
        for (const [id, at] of offsets) {
          const v = s.reader(at, ends.get(at) - at, 'OLE property ' + id);
          // Office's HeadingPairs/DocumentParts LPSTR elements have no padding. Other strings
          // permit missing padding too, as POI does, but never eat a nonzero next byte.
          const unaligned = officeSummary && (id === 12 || id === 13);
          const pad = p => { const end = Math.ceil(p / 4) * 4; while (p < end && p < v.length && v.u8(p) === 0) p++; return p; };
          let budget = 100000;
          const read = (at, forced, depth = 0) => {
            if (--budget < 0 || depth > 16) B.fail('PROPERTIES_LIMIT', N('Property vector is too large or nested'));
            const type = forced ?? v.u16(at); let p = at + (forced == null ? 4 : 0), value, ticks, children;
            if (type & 0x1000) {
              const len = v.u32(p); p += 4; if (len > budget) B.fail('PROPERTIES_LIMIT', N('Property vector is too large'));
              children = []; value = [];
              for (let k = 0; k < len; k++) { const c = read(p, (type & 0xfff) === 12 ? undefined : type & 0xfff, depth + 1); children.push(c); value.push(c.value); p = c.end; }
            } else switch (type) {
              case 0: case 1: value = null; break;
              case 16: value = v.u8(p); if (value > 127) value -= 256; p++; break;
              case 17: value = v.u8(p++); break;
              case 2: value = v.i16(p); p += 2; break;
              case 18: value = v.u16(p); p += 2; break;
              case 3: case 22: value = v.i32(p); p += 4; break;
              case 10: case 19: case 23: value = v.u32(p); p += 4; break;
              case 4: value = v.f32(p); p += 4; break;
              case 5: value = v.f64(p); p += 8; break;
              case 11: value = !!v.u16(p); p += 2; break;
              case 6: case 20: case 21: {
                let number = BigInt(v.u32(p)) | BigInt(v.u32(p + 4)) << 32n; p += 8;
                if (type !== 21) number = BigInt.asIntN(64, number);
                value = type === 6 ? (number < 0n ? '-' : '') + (number < 0n ? -number : number) / 10000n + '.' + String((number < 0n ? -number : number) % 10000n).padStart(4, '0') : String(number); break;
              }
              case 7: { const days = v.f64(p); p += 8; value = date(Date.UTC(1899, 11, 30) + (Math.trunc(days) + Math.abs(days % 1)) * 86400000); break; }
              case 8: case 30: case 31: {
                const len = v.u32(p); p += 4; const bytes = len * (type === 31 ? 2 : 1);
                if (bytes > B.limits.text) B.fail('PROPERTIES_LIMIT', N('Property string is too large'));
                value = v.text(p, bytes, type === 31 ? 1200 : set.codepage).replace(/\0+$/, ''); p += bytes;
                if (!(unaligned && type === 30)) p = pad(p); break;
              }
              case 64: {
                const t = BigInt(v.u32(p)) | BigInt(v.u32(p + 4)) << 32n; ticks = String(t); p += 8;
                value = t === 0n ? null : date(Number(t / 10000n) - 11644473600000); break;
              }
              case 65: case 70: case 71: {
                const len = v.u32(p); p += 4; value = [...v.sub(p, len)]; p = pad(p + len); break;
              }
              case 72: value = guid(v, p); p += 16; break;
              default: B.fail('PROPERTY_TYPE', N('This document property has an unsupported type'), type);
            }
            if (forced == null && !(unaligned && type === 30)) p = pad(p);
            return { type, value, end: p, ...(ticks == null ? {} : { ticks }), ...(children ? { children } : {}) };
          };
          try {
            if (id === 0) {
              const len = v.u32(0); if (len > 100000) B.fail('PROPERTIES_LIMIT', N('Property dictionary is too large'));
              let p = 4;
              for (let k = 0; k < len; k++) {
                const key = v.u32(p), chars = v.u32(p + 4), width = set.codepage === 1200 ? 2 : 1; p += 8;
                dictionary[key] = v.text(p, chars * width, set.codepage).replace(/\0+$/, ''); p += chars * width;
                if (width === 2) p = pad(p);
              }
            } else values[id] = read(0);
          } catch (e) {
            if (e.name !== 'LegacyError') throw e;
            issues.push({ id, code: e.code }); values[id] = { type: id === 0 ? null : v.u16(0), value: null, unsupported: true, error: e.code };
          }
        }
      } catch (e) { if (e.name !== 'LegacyError') throw e; issues.push({ code: e.code }); }
    }
    return sets;
  };
})(typeof window !== 'undefined' ? window : globalThis);
