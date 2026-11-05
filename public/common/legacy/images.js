/* Bounded image primitives for OfficeArt: MS-WMF DIB and PNG (ISO/IEC 15948).
 * PNG uses the browser's DEFLATE compressor, with stored blocks as a fallback.
 * CRC32 comes from the shared ZIP library. */
(function (root) {
  'use strict';
  const N = s => s; // UI refusal text, translated where it is shown.
  const L = root.L, B = L.legacy, I = B.images = {};
  const maxImage = 64 * 1024 * 1024;
  I.inflate = async function (bytes, expected, limit = maxImage) {
    if (!Number.isSafeInteger(expected) || expected < 0 || expected > limit) B.fail('IMAGE_LIMIT', N('The decompressed picture is too large'));
    if (typeof DecompressionStream === 'undefined') B.fail('DECOMPRESSION_UNAVAILABLE', N('This browser cannot decompress metafiles'));
    const reader = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate')).getReader();
    const chunks = []; let size = 0;
    try {
      for (;;) { const { value, done } = await reader.read(); if (done) break; size += value.length;
        if (size > limit || size > expected) { await reader.cancel(); B.fail('IMAGE_LIMIT', N('The decompressed picture exceeds its declared size')); } chunks.push(value); }
    } catch (e) { if (e.name === 'LegacyError') throw e; B.fail('IMAGE_COMPRESSION', N('The picture contains invalid compressed data')); }
    finally { reader.releaseLock(); }
    if (size !== expected) B.fail('IMAGE_SIZE', N('The decompressed picture has the wrong size'));
    const out = new Uint8Array(size); let p = 0; for (const c of chunks) { out.set(c, p); p += c.length; } return out;
  };
  const join = arrays => { const size = arrays.reduce((s, b) => s + b.length, 0); if (size > B.limits.bytes) B.fail('IMAGE_LIMIT', N('The picture is too large')); const out = new Uint8Array(size); let p = 0; for (const b of arrays) { out.set(b, p); p += b.length; } return out; };
  const be32 = n => { const b = new Uint8Array(4); new DataView(b.buffer).setUint32(0, n); return b; };
  function chunk(type, data) { const body = join([L.zip.encode(type), data]); return join([be32(data.length), body, be32(L.zip.crc32(body))]); }
  async function deflate(raw) {
    if (typeof CompressionStream === 'function') {
      try {
        const stream = new Blob([raw]).stream().pipeThrough(new CompressionStream('deflate'));
        return new Uint8Array(await new Response(stream).arrayBuffer());
      } catch { /* Older engines can expose the API without supporting this format. */ }
    }
    const z = [new Uint8Array([0x78, 1])]; let a = 1, b = 0;
    for (let p = 0; p < raw.length; p += 65535) {
      const len = Math.min(65535, raw.length - p), h = new Uint8Array([p + len === raw.length ? 1 : 0, len & 255, len >> 8, ~len & 255, ~len >> 8 & 255]);
      z.push(h, raw.subarray(p, p + len));
      for (let j = p; j < p + len; j++) { a = (a + raw[j]) % 65521; b = (b + a) % 65521; }
    }
    z.push(be32((b * 65536 + a) >>> 0));
    return join(z);
  }
  I.png = async function (width, height, rgba) {
    if (!Number.isInteger(width) || !Number.isInteger(height) || width <= 0 || height <= 0 || height * (width * 4 + 1) > maxImage || rgba.length !== width * height * 4) B.fail('IMAGE_LIMIT', N('Invalid picture dimensions'));
    const raw = new Uint8Array(height * (width * 4 + 1));
    for (let y = 0; y < height; y++) raw.set(rgba.subarray(y * width * 4, (y + 1) * width * 4), y * (width * 4 + 1) + 1);
    const compressed = await deflate(raw);
    const ihdr = join([be32(width), be32(height), new Uint8Array([8, 6, 0, 0, 0])]);
    return join([new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]), chunk('IHDR', ihdr), chunk('IDAT', compressed), chunk('IEND', new Uint8Array(0))]);
  };
  I.dib = async function (bytes) {
    let r = new B.Reader(bytes, 'DIB'), fileOffset;
    if (r.length >= 14 && r.u16(0) === 0x4d42) { fileOffset = r.u32(10) - 14; r = r.reader(14, r.length - 14); }
    const header = r.u32(0), core = header === 12;
    if (![12, 40, 52, 56, 108, 124].includes(header)) B.fail('DIB_HEADER', N('Unsupported bitmap header'));
    r.check(0, header);
    const width = core ? r.u16(4) : r.i32(4), signedHeight = core ? r.u16(6) : r.i32(8), height = Math.abs(signedHeight), top = signedHeight < 0;
    const planes = r.u16(core ? 8 : 12), bits = r.u16(core ? 10 : 14), compression = core ? 0 : r.u32(16);
    if (planes !== 1 || width <= 0 || height <= 0 || height * (width * 4 + 1) > maxImage) B.fail('IMAGE_LIMIT', N('Invalid or oversized bitmap dimensions'));
    if (![1, 4, 8, 16, 24, 32].includes(bits) || ![0, 1, 2, 3, 6].includes(compression) || compression === 1 && bits !== 8 || compression === 2 && bits !== 4 || [3, 6].includes(compression) && ![16, 32].includes(bits)) B.fail('DIB_ENCODING', N('Unsupported bitmap encoding'));
    let offset = header, masks = bits === 16 ? [0x7c00, 0x3e0, 0x1f, 0] : [0xff0000, 0xff00, 0xff, 0];
    if ([3, 6].includes(compression)) {
      const count = header >= 56 || compression === 6 ? 4 : 3, p = header >= 52 ? 40 : header;
      masks = Array.from({ length: 4 }, (_, i) => i < count ? r.u32(p + i * 4) : 0); if (header === 40) offset += count * 4;
    }
    const channels = masks.map((mask, i) => {
      if (!mask) return () => i === 3 ? 255 : 0;
      let shift = 0; while (((mask >>> shift) & 1) === 0) shift++;
      const max = mask >>> shift;
      if ((max & (max + 1)) !== 0) B.fail('DIB_MASK', N('Unsupported noncontiguous bitmap channel'));
      return pixel => Math.round(((pixel & mask) >>> shift) * 255 / max);
    });
    const colours = bits <= 8 ? (core ? 0 : r.u32(32)) || 2 ** bits : core ? 0 : r.u32(32);
    if (colours > 256 || bits <= 8 && colours > 2 ** bits) B.fail('DIB_PALETTE', N('Invalid bitmap palette'));
    const palette = [], entrySize = core ? 3 : 4;
    r.check(offset, colours * entrySize);
    for (let i = 0; i < colours; i++) { const p = offset + i * entrySize; palette.push([r.u8(p + 2), r.u8(p + 1), r.u8(p), 255]); }
    offset += colours * entrySize;
    if (fileOffset != null) { if (fileOffset < offset) B.fail('DIB_OFFSET', N('Bitmap pixels overlap its header')); offset = fileOffset; }
    const rgba = new Uint8Array(width * height * 4), set = (x, y, colour) => {
      if (x < 0 || x >= width || y < 0 || y >= height || !colour) B.fail('DIB_PIXELS', N('Invalid bitmap pixel or palette index'));
      rgba.set(colour, ((top ? y : height - y - 1) * width + x) * 4);
    };
    if (compression === 1 || compression === 2) {
      if (top) B.fail('DIB_ENCODING', N('Top-down RLE bitmaps are not supported'));
      // Unwritten RLE pixels use palette index0 (including skipped rows and deltas).
      for (let i = 0; i < rgba.length; i += 4) rgba.set(palette[0], i);
      let p = offset, x = 0, y = 0, end = false;
      while (p < r.length && !end) {
        const count = r.u8(p++), value = r.u8(p++);
        if (count) { for (let i = 0; i < count; i++) set(x++, y, palette[bits === 8 ? value : i % 2 ? value & 15 : value >> 4]); }
        else if (value === 0) { x = 0; y++; }
        else if (value === 1) end = true;
        else if (value === 2) { x += r.u8(p++); y += r.u8(p++); if (x > width || y >= height) B.fail('DIB_PIXELS', N('Invalid bitmap RLE delta')); }
        else {
          const len = bits === 8 ? value : Math.ceil(value / 2); r.check(p, len);
          for (let i = 0; i < value; i++) { const v = r.u8(p + (bits === 8 ? i : i >> 1)); set(x++, y, palette[bits === 8 ? v : i % 2 ? v & 15 : v >> 4]); }
          p += len + (len & 1); r.check(p, 0);
        }
      }
      if (!end) B.fail('DIB_PIXELS', N('The bitmap RLE stream is truncated'));
    } else {
      const stride = Math.ceil(width * bits / 32) * 4; r.check(offset, stride * height);
      for (let y = 0; y < height; y++) for (let x = 0; x < width; x++) {
        const p = offset + y * stride + Math.floor(x * bits / 8); let colour;
        if (bits <= 8) { const n = r.u8(p), index = bits === 8 ? n : bits === 4 ? x % 2 ? n & 15 : n >> 4 : n >> (7 - (x & 7)) & 1; colour = palette[index]; }
        else if (bits === 24) colour = [r.u8(p + 2), r.u8(p + 1), r.u8(p), 255];
        else { const pixel = bits === 16 ? r.u16(p) : r.u32(p); colour = channels.map(c => c(pixel)); }
        set(x, y, colour);
      }
    }
    return { bytes: await I.png(width, height, rgba), ext: 'png', mime: 'image/png', width, height };
  };
})(typeof window !== 'undefined' ? window : globalThis);
