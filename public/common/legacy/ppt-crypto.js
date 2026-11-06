/* PowerPoint CryptoAPI encryption: MS-PPT 2.3.7 and MS-ODRAW 2.2.22.
 * Derived from Apache POI hslf/usermodel/HSLFSlideShowEncrypted.java (Apache-2.0).
 * A persist object uses its ID as the block number. Picture fields each restart block 0. */
(function (root) {
  'use strict';
  const B = root.L.legacy, P = B.ppt, E = B.encryption, N = s => s;
  const ENCRYPTED_USER = 0xf3d1c4df, CRYPT_SESSION = 0x2f14, DOCUMENT = 1000;
  const BSE = 0xf007, EMF = 0xf01a, WMF = 0xf01b, PICT = 0xf01c;
  const TWO_UIDS = new Set([0x217, 0x3d5, 0x46b, 0x543, 0x6e1, 0x6e3, 0x6e5, 0x7a9]);
  const BSE_FIELDS = [1, 1, 16, 2, 4, 4, 4, 1, 1, 1, 1];
  P.decryptPictures = function (bytes, session) {
    const out = bytes.slice(), r = new B.Reader(out, 'Pictures');
    const field = (at, n, end) => {
      r.check(at, n);
      if (at + n > end) B.fail('PPT_PICTURE', N('An encrypted picture is damaged'));
      out.set(E.rc4(session.key(0))(bytes.subarray(at, at + n)), at);
    };
    let at = 0, count = 0;
    while (at < r.length) {
      if (++count > B.limits.records) B.fail('RECORD_LIMIT', N('Too many presentation pictures'));
      let end;
      try {
        field(at, 8, r.length);
        end = at + 8 + r.u32(at + 4);
        r.check(at, end - at);
        let type = r.u16(at + 2), instance = r.u16(at) >>> 4, p = at + 8;
        if (type === BSE) {
          for (const size of BSE_FIELDS) {
            field(p, size, end);
            p += size;
          }
          const nameLength = r.u8(at + 8 + 33);
          field(p, nameLength, end);
          p += nameLength;
          if (p === end) {
            at = end;
            continue;
          }
          field(p, 8, end);
          type = r.u16(p + 2);
          instance = r.u16(p) >>> 4;
          p += 8;
        }
        const uids = TWO_UIDS.has(instance) ? 2 : 1;
        for (let i = 0; i < uids; i++) {
          field(p, 16, end);
          p += 16;
        }
        const size = [EMF, WMF, PICT].includes(type) ? 34 : 1;
        field(p, size, end);
        p += size;
        field(p, end - p, end);
      } catch (e) {
        if (e.name !== 'LegacyError') throw e;
        // Keep good earlier/later pictures. Their BSE offsets still address the
        // original stream; the normal BLIP reader reports this one locally.
        out.fill(0, at, Math.min(end > at ? end : r.length, r.length));
        if (!(end > at && end <= r.length)) break;
      }
      at = end;
    }
    return out;
  };
  P.decrypt = function (cfb, options = {}) {
    const userBytes = cfb.stream('Current User');
    if (!userBytes || userBytes.length < 28) return cfb;
    const user = new B.Reader(userBytes, 'Current User');
    const encrypted = user.u32(12) === ENCRYPTED_USER;
    // A few producers leave a plain CurrentUser token on encrypted files. The
    // optional UserEdit field, not just that token, identifies the encryption atom.
    const bytes = cfb.stream('PowerPoint Document');
    if (!bytes) B.fail('PPT_STREAM', N('The PowerPoint Document stream is missing'));
    const r = new B.Reader(bytes, 'PowerPoint Document'), directory = P.directory(r, user);
    const offset = directory.persists.get(directory.cryptRef);
    if (offset == null) {
      if (!encrypted) return cfb;
      // POI also tolerates stale encryption references from SoftMaker if the
      // current document is really plaintext. Do not interpret ciphertext as text.
      const document = P.record(r, directory.persists.get(directory.docRef));
      if (document.type !== DOCUMENT) {
        B.fail('ENCRYPTION_HEADER', N('The presentation password information is missing'));
      }
      return { ...cfb, pptDirectory: directory, pptRecord: (id, at) => P.record(r, at) };
    }
    const atom = P.record(r, offset);
    if (atom.type !== CRYPT_SESSION) {
      B.fail('ENCRYPTION_HEADER', N('The presentation password information is damaged'));
    }
    const session = E.unlock(E.header(atom.r.bytes), options.password, 'Presentation');
    const records = new Map();
    let decodedBytes = 0;
    const view = E.overlay(cfb, (name, key) => {
      const data = cfb.stream(name);
      return key === 'pictures' && data ? P.decryptPictures(data, session) : data;
    });
    return { ...E.summary(view, session, 'EncryptedSummary'), pptDirectory: directory, pptRecord(id, at) {
      if (!records.has(id)) {
        const cipher = E.rc4(session.key(id)), head = cipher(r.sub(at, 8));
        const size = new B.Reader(head).u32(4);
        if (size + 8 > B.limits.bytes - decodedBytes) {
          B.fail('SIZE_LIMIT', N('The decrypted presentation is too large to import'));
        }
        const data = E.join(head, cipher(r.sub(at + 8, size)));
        const record = P.record(new B.Reader(data, 'Persist ' + id), 0);
        record.offset = at;
        record.next = at + data.length;
        records.set(id, record);
        decodedBytes += data.length;
      }
      return records.get(id);
    } };
  };
})(typeof window !== 'undefined' ? window : globalThis);
