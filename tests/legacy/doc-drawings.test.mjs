import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import { load } from '../../tools/legacy/load.mjs';
import { doc, u16, u32, sprm } from './doc-fixtures.mjs';
import { rec, container } from './ppt-fixtures.mjs';
const B = load(),
  D = B.doc;
if (process.env.LEGACY_BEFORE_E3)
  for (const name of [
    'officeart',
    'shapes',
    'doc-read',
    'doc-sprm',
    'doc-structure',
    'doc-style',
    'doc-lists',
    'doc-xml',
    'doc-settings',
    'doc-fields',
    'doc-markup',
    'doc-table-props',
    'doc-tables',
    'doc-text',
    'doc-write'
  ])
    vm.runInThisContext(fs.readFileSync(process.env.LEGACY_BEFORE_E3 + '/' + name + '.js', 'utf8'));
if (process.env.LEGACY_BEFORE_DRAW_LOCAL)
  for (const name of ['officeart', 'doc-draw-read', 'doc-draw-write'])
    vm.runInThisContext(fs.readFileSync(process.env.LEGACY_BEFORE_DRAW_LOCAL + '/' + name + '.js', 'utf8'));
if (process.env.LEGACY_BEFORE_DRAW_NOTICES)
  for (const name of ['doc-draw-read', 'doc-draw-write', 'doc-text'])
    vm.runInThisContext(fs.readFileSync(process.env.LEGACY_BEFORE_DRAW_NOTICES + '/' + name + '.js', 'utf8'));
const join = (...parts) => Buffer.concat(parts),
  png = Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aXioAAAAASUVORK5CYII=',
    'base64'
  );
const blip = () => rec(0xf01e, join(Buffer.alloc(16), Buffer.from([255]), png), 0x6e0);
const bse = (data = blip()) => {
  const b = Buffer.alloc(36);
  b[0] = b[1] = 6;
  b.writeUInt32LE(data.length, 20);
  b.writeUInt32LE(1, 24);
  b.writeUInt32LE(0xffffffff, 28);
  return rec(0xf007, join(b, data), 6, 2);
};
const opt = props =>
  rec(
    0xf00b,
    join(
      ...props.map(([id, v]) =>
        join(u16(id | (Buffer.isBuffer(v) ? 0x8000 : 0)), u32(Buffer.isBuffer(v) ? v.length : v))
      ),
      ...props.filter(p => Buffer.isBuffer(p[1])).map(p => p[1])
    ),
    props.length,
    3
  );
const shape = (id, type = 75, props = [], extra = [], flags = 0x0a00) =>
  container(0xf004, [rec(0xf00a, join(u32(id), u32(flags)), type, 2), opt(props), ...extra]);
const drawing = (shapes, entries = [], label = 0) =>
  join(
    container(0xf000, entries.length ? [container(0xf001, entries)] : []),
    Buffer.from([label]),
    container(0xf002, shapes)
  );
const spa = (anchors, end) =>
  join(
    ...anchors.map(a => u32(a.cp)),
    u32(end),
    ...anchors.map(a =>
      join(
        u32(a.id),
        ...(a.rect || [100, 200, 1540, 920]).map(u32),
        u16(a.flags ?? (2 << 5) | (2 << 1) | (2 << 3)),
        u32(0)
      )
    )
  );
const txbx = (text, id = 1025) => {
  const b = Buffer.alloc(22);
  b.writeUInt32LE(id, 14);
  return join(u32(0), u32(text.length), b);
};
const picture = (props = [], mm = 100) => {
  const parts = join(shape(1025, 75, [[260, 2], ...props]), bse()),
    h = Buffer.alloc(68);
  h.writeUInt32LE(68 + (mm === 102 ? 5 : 0) + parts.length, 0);
  h.writeUInt16LE(68, 4);
  h.writeUInt16LE(mm, 6);
  h.writeInt16LE(1440, 28);
  h.writeInt16LE(720, 30);
  h.writeUInt16LE(1000, 32);
  h.writeUInt16LE(1000, 34);
  return join(h, ...(mm === 102 ? [Buffer.from([4]), Buffer.from('name')] : []), parts);
};
function inlineDoc(data = picture(), text = 'Before \x01 after\r', extra = {}) {
  const cp = text.indexOf('\x01');
  return doc({
    text,
    data,
    chpx: [
      { start: 2048, end: 2048 + cp * 2, ops: [] },
      { start: 2048 + cp * 2, end: 2048 + (cp + 1) * 2, ops: [sprm(0x6a03, 0), sprm(0x0855, 1)] },
      { start: 2048 + (cp + 1) * 2, end: 2048 + text.length * 2, ops: [] }
    ],
    ...extra
  });
}
async function convert(bytes) {
  const r = await D.convert(bytes),
    zip = await L.zip.read(r.pkg);
  return { ...r, zip, xml: await zip.get('word/document.xml').text() };
}
const mediaNames = r => [...r.zip.keys()].filter(n => /^word\/media\//.test(n));
const texts = xml => {
  let out = '';
  const visit = () => {};
  visit.text = (s, n) => {
    if (n?.name === 'w:t') out += s;
  };
  B.checkXml(xml, visit);
  return out;
};

test('Word inline PICF writes its private PNG, size, crop and alternative text', async () => {
  const r = await convert(
    inlineDoc(
      picture(
        [
          [258, 16384],
          [897, Buffer.from('A <picture>\0', 'utf16le')]
        ],
        102
      )
    )
  );
  assert.match(r.xml, /<wp:inline/);
  assert.match(r.xml, /<wp:extent cx="914400" cy="457200"/);
  assert.match(r.xml, /<a:srcRect l="25000"/);
  assert.match(r.xml, /descr="A &lt;picture&gt;"/);
  assert.equal(texts(r.xml), 'Before  after');
  assert.equal(r.stats.pictures, 1);
  assert.equal(mediaNames(r).length, 1);
  assert.deepEqual(Buffer.from(await r.zip.get(mediaNames(r)[0]).bytes()), png);
  assert.ok(!r.losses.some(x => x.id === 'legacy-word-picture'));
  assert.match(await r.zip.get('[Content_Types].xml').text(), /<Default Extension="png" ContentType="image\/png"/);
});
test('Word damaged optional picture keeps its surrounding text', async () => {
  const data = picture();
  data.writeUInt32LE(0x7fffffff, 0);
  const r = await convert(inlineDoc(data));
  assert.equal(texts(r.xml), 'Before  after');
  assert.equal(mediaNames(r).length, 0);
  assert.ok(r.losses.some(x => x.notify && /picture/i.test(x.what)));
});
test('Word floating anchors preserve placement, wrap, rotation and reflection', async () => {
  const text = '\x08Body\r',
    s = shape(
      1031,
      75,
      [
        [260, 1],
        [4, 90 * 65536],
        [912, 1],
        [914, 2]
      ],
      [],
      0xa40
    );
  const r = await convert(
    doc({
      text,
      tables: [
        [50, drawing([s], [bse()])],
        [40, spa([{ id: 1031, cp: 0, flags: 1 << 5, rect: [100, 200, 820, 1640] }], text.length)]
      ]
    })
  );
  assert.match(r.xml, /<wp:positionH relativeFrom="page"><wp:posOffset>-165100/);
  assert.match(r.xml, /<wp:positionV relativeFrom="paragraph">/);
  assert.match(r.xml, /<wp:wrapTopAndBottom\/>/);
  assert.match(r.xml, /<wp:extent cx="914400" cy="457200"/);
  assert.match(r.xml, /<a:xfrm rot="5400000" flipH="1"><a:off x="0" y="0"/);
  assert.equal(r.stats.pictures, 1);
});
test('Word drawings in headers own their image relationships even when the header drawing comes first', async () => {
  const text = '\x08Body\r',
    header = '\x08Header\r\r',
    s = shape(1025, 75, [[260, 1]]),
    h = shape(2049, 75, [[260, 1]]);
  const data = join(drawing([h], [bse()], 1), Buffer.from([0]), container(0xf002, [s]));
  const r = await convert(
    doc({
      text,
      stories: { headers: header },
      headers: [0, 0, 0, 0, 0, 0, 0, header.length],
      tables: [
        [50, data],
        [40, spa([{ id: 1025, cp: 0 }], text.length)],
        [41, spa([{ id: 2049, cp: 0 }], header.length)]
      ]
    })
  );
  assert.match(await r.zip.get('word/header1.xml').text(), /<wp:anchor/);
  assert.match(r.xml, /<wp:anchor/);
  assert.equal(mediaNames(r).length, 1);
  for (const name of ['document', 'header1'])
    assert.match(await r.zip.get('word/_rels/' + name + '.xml.rels').text(), /Target="media\/image1.png"/);
});
test('Word a missing global image index does not substitute an unrelated picture', async () => {
  const text = '\x08\x08Body\r',
    r = await convert(
      doc({
        text,
        tables: [
          [50, drawing([shape(1025, 75, [[260, 99]]), shape(1026, 75, [[260, 1]])], [bse()])],
          [
            40,
            spa(
              [
                { id: 1025, cp: 0 },
                { id: 1026, cp: 1 }
              ],
              text.length
            )
          ]
        ]
      })
    );
  assert.equal(r.stats.pictures, 1);
  assert.equal((r.xml.match(/<wp:anchor /g) || []).length, 1);
  assert.ok(r.losses.some(x => x.notify && /picture/i.test(x.what)));
});
test('Word a bad Data fallback is never read when WordDocument has the delayed image', async () => {
  const bytes = join(Buffer.alloc(16), blip()),
    entry = bse(Buffer.alloc(0));
  entry.writeUInt32LE(16, 8 + 28);
  let reads = 0;
  const store = B.officeart.store(B.officeart.records(container(0xf000, [container(0xf001, [entry])])), [
    bytes,
    () => {
      reads++;
      throw Object.assign(Error('broken Data'), { name: 'CFBError' });
    }
  ]);
  const image = await store[0].read();
  assert.equal(image.ext, 'png');
  assert.equal(reads, 0);
});
test('Word delayed images can use a lazy Data stream after the primary stream fails', async () => {
  const bytes = join(Buffer.alloc(16), blip()),
    entry = bse(Buffer.alloc(0));
  entry.writeUInt32LE(16, 8 + 28);
  let reads = 0;
  const store = B.officeart.store(B.officeart.records(container(0xf000, [container(0xf001, [entry])])), [
    Buffer.alloc(16),
    () => {
      reads++;
      return bytes;
    }
  ]);
  assert.equal((await store[0].read()).ext, 'png');
  assert.equal((await store[0].read()).ext, 'png');
  assert.equal(reads, 1);
});
test('Word native text boxes retain paragraphs, formatting and interior margins', async () => {
  const text = '\x08Body\r',
    inside = 'Box text\r',
    s = shape(
      1025,
      202,
      [
        [129, 457200],
        [130, 914400],
        [385, 0xd8d8d8],
        [447, 0x600000]
      ],
      [rec(0xf00d, u32(65536))]
    );
  const r = await convert(
    doc({
      text,
      stories: { textboxes: inside },
      tables: [
        [50, drawing([s])],
        [40, spa([{ id: 1025, cp: 0 }], text.length)],
        [56, txbx(inside)]
      ]
    })
  );
  assert.match(r.xml, /<v:textbox inset="36pt,72pt,7.2pt,3.6pt"><w:txbxContent>/);
  assert.equal(texts(r.xml), 'Box textBody');
  assert.equal(r.stats.textboxes, 1);
  assert.ok(!r.losses.some(x => x.notify));
  assert.match(r.xml, /filled="t"[^>]*fillcolor="#D8D8D8"/);
});
test('Word linked text boxes write one live story with a native next-box reference', async () => {
  const text = '\x08\x08Body\r',
    inside = 'One story\r',
    shapes = [shape(1025, 202, [], [rec(0xf00d, u32(65536))]), shape(1026, 202, [], [rec(0xf00d, u32(65537))])];
  const r = await convert(
    doc({
      text,
      stories: { textboxes: inside },
      tables: [
        [50, drawing(shapes)],
        [
          40,
          spa(
            [
              { id: 1025, cp: 0 },
              { id: 1026, cp: 1 }
            ],
            text.length
          )
        ],
        [56, txbx(inside)]
      ]
    })
  );
  assert.match(r.xml, /mso-next-textbox:#_x0000_s1026/);
  assert.equal(texts(r.xml), 'One storyBody');
  assert.equal(r.stats.textboxes, 2);
});
test('Word grouped shapes retain their child coordinates and live text boxes', async () => {
  const text = '\x08Body\r',
    inside = 'Inside group\r',
    child = shape(1026, 202, [], [rec(0xf00f, join(...[10, 20, 90, 80].map(u32))), rec(0xf00d, u32(65536))]);
  const group = container(0xf003, [
    shape(1025, 0, [], [rec(0xf009, join(...[0, 0, 100, 100].map(u32)))], 0xa01),
    child
  ]);
  const r = await convert(
    doc({
      text,
      stories: { textboxes: inside },
      tables: [
        [50, drawing([group])],
        [40, spa([{ id: 1025, cp: 0 }], text.length)],
        [56, txbx(inside, 1026)]
      ]
    })
  );
  assert.match(r.xml, /<v:group[^>]*coordorigin="0,0" coordsize="100,100"/);
  assert.match(r.xml, /left:10;top:20;width:80;height:60/);
  assert.equal(texts(r.xml), 'Inside groupBody');
  assert.equal(r.stats.groups, 1);
});
test('Word unsupported text-box shape geometry keeps the text in a plain box', async () => {
  const text = '\x08Body\r',
    inside = 'Readable\r',
    s = shape(1025, 300, [], [rec(0xf00d, u32(65536))]);
  const r = await convert(
    doc({
      text,
      stories: { textboxes: inside },
      tables: [
        [50, drawing([s])],
        [40, spa([{ id: 1025, cp: 0 }], text.length)],
        [56, txbx(inside)]
      ]
    })
  );
  assert.equal(texts(r.xml), 'ReadableBody');
  assert.match(r.xml, /<v:rect /);
  assert.ok(r.losses.some(x => x.notify && x.id === 'legacy-word-drawing-style'));
});
test('Word brightness and contrast use native VML image adjustments', async () => {
  const r = await convert(
    inlineDoc(
      picture([
        [264, 32768],
        [265, 16384],
        [319, 0x60004]
      ])
    )
  );
  assert.match(r.xml, /<v:imagedata[^>]*gain="0.5" blacklevel="0.25" grayscale="t"/);
  assert.doesNotMatch(r.xml, /bilevel="t"|position:absolute/);
  assert.ok(!r.losses.some(x => x.notify));
});
test('Word a drawing at the paragraph mark is emitted without eating text', async () => {
  const text = 'Body\r',
    r = await convert(
      doc({
        text,
        tables: [
          [50, drawing([shape(1025, 1)])],
          [40, spa([{ id: 1025, cp: 4 }], text.length)]
        ]
      })
    );
  assert.equal(texts(r.xml), 'Body');
  assert.match(r.xml, /<v:rect /);
  assert.ok(!r.losses.some(x => x.notify));
});
test('Word paragraph frames decode the high anchor bits and signed alignment values', async () => {
  const text = 'Frame\r',
    r = await convert(
      doc({
        text,
        papx: [
          {
            start: 2048,
            end: 2048 + text.length * 2,
            ops: [sprm(0x261b, 0x50), sprm(0x8418, 65532), sprm(0x8419, 65528), sprm(0x841a, 1440)]
          }
        ]
      })
    );
  assert.match(
    r.xml,
    /<w:framePr[^>]*w:vAnchor="page" w:hAnchor="margin"[^>]*w:xAlign="center"[^>]*w:yAlign="center"/
  );
});
test('Word frames preserve plus-one offsets, zero alignments and wrap kind', async () => {
  const text = 'First\rSecond\r',
    r = await convert(
      doc({
        text,
        papx: [
          { start: 2048, end: 2060, ops: [sprm(0x8418, 21), sprm(0x8419, 41), sprm(0x2423, 3)] },
          { start: 2060, end: 2048 + text.length * 2, ops: [sprm(0x8418, 0), sprm(0x8419, 0)] }
        ]
      })
    );
  assert.match(r.xml, /<w:framePr[^>]*w:wrap="none"[^>]*w:x="20"[^>]*w:y="40"/);
  assert.match(r.xml, /<w:framePr[^>]*w:xAlign="left"[^>]*w:yAlign="inline"/);
});
test('Word drawing defaults merge Boolean flags individually with shape overrides', async () => {
  const text = '\x08Body\r',
    s = shape(1025, 1, [[447, 0x200020]]),
    art = join(
      container(0xf000, [
        opt([
          [385, 0x563412],
          [447, 0x100000]
        ])
      ]),
      Buffer.from([0]),
      container(0xf002, [s])
    );
  const r = await convert(
    doc({
      text,
      tables: [
        [50, art],
        [40, spa([{ id: 1025, cp: 0 }], text.length)]
      ]
    })
  );
  assert.match(r.xml, /filled="f"[^>]*fillcolor="#123456"/);
});
test('Word lost live text boxes notify while reusable slots do not resurrect deleted text', async () => {
  const bytes = txbx('Lost text\r'),
    r = await convert(doc({ text: 'Body\r', stories: { textboxes: 'Lost text\r' }, tables: [[56, bytes]] }));
  assert.equal(texts(r.xml), 'Body');
  assert.ok(r.losses.some(x => x.notify && x.id === 'legacy-word-textboxes'));
  bytes.writeUInt16LE(1, 8 + 8);
  const reusable = await convert(
    doc({ text: 'Body\r', stories: { textboxes: 'Lost text\r' }, tables: [[56, bytes]] })
  );
  assert.ok(!reusable.losses.some(x => x.id === 'legacy-word-textboxes'));
});
test('Word unsupported embedded-object markers are reported locally', async () => {
  const text = 'Before \x01 after\r',
    cp = text.indexOf('\x01'),
    r = await convert(
      doc({
        text,
        chpx: [
          { start: 2048, end: 2048 + cp * 2, ops: [] },
          { start: 2048 + cp * 2, end: 2048 + (cp + 1) * 2, ops: [sprm(0x0806, 1)] },
          { start: 2048 + (cp + 1) * 2, end: 2048 + text.length * 2, ops: [] }
        ]
      })
    );
  assert.equal(texts(r.xml), 'Before  after');
  assert.ok(r.losses.some(x => x.notify && /object|picture/i.test(x.what)));
});
test('Word HYPERLINK instruction data is not a lost embedded object', async () => {
  const text = '\x13 HYPERLINK "https://example.test/"\x01\x14Link\x15\r',
    cp = text.indexOf('\x01'),
    r = await convert(
      doc({
        text,
        chpx: [
          { start: 2048, end: 2048 + cp * 2, ops: [] },
          { start: 2048 + cp * 2, end: 2048 + (cp + 1) * 2, ops: [sprm(0x0806, 1)] },
          { start: 2048 + (cp + 1) * 2, end: 2048 + text.length * 2, ops: [] }
        ]
      })
    );
  assert.equal(texts(r.xml), 'Link');
  assert.ok(!r.losses.some(x => x.notify));
  assert.match(r.xml, /HYPERLINK &quot;https:\/\/example.test\//);
});
test('Word reference and localized field metadata does not produce content-loss notices', async () => {
  for (const keyword of ['PAGEREF', 'REF', 'RENVOIPAGE', 'PRIVATE']) {
    const text = '\x13 ' + keyword + ' example \\h \x01\x14Result\x15\r',
      cp = text.indexOf('\x01'),
      r = await convert(
        doc({
          text,
          chpx: [
            { start: 2048, end: 2048 + cp * 2, ops: [] },
            { start: 2048 + cp * 2, end: 2048 + (cp + 1) * 2, ops: [sprm(0x0806, 1)] },
            { start: 2048 + (cp + 1) * 2, end: 2048 + text.length * 2, ops: [] }
          ]
        })
      );
    assert.equal(texts(r.xml), 'Result');
    assert.ok(!r.losses.some(x => x.notify), keyword);
  }
});
test('Word floating OLE preview keeps its image without a false missing-text-box notice', async () => {
  const text = '\x08Body\r',
    inside = '\x13 EMBED Equation.3 \x14\x01\x15\r\r',
    r = await convert(
      doc({
        text,
        stories: { textboxes: inside },
        tables: [
          [50, drawing([shape(1025, 75, [[260, 1]])], [bse()])],
          [40, spa([{ id: 1025, cp: 0 }], text.length)],
          [56, txbx(inside)]
        ]
      })
    );
  assert.equal(r.stats.pictures, 1);
  assert.ok(r.losses.some(x => x.notify && x.what === 'An embedded object was kept only as a picture'));
  assert.ok(!r.losses.some(x => x.id === 'legacy-word-textboxes' || x.what === 'A picture could not be read'));
});
test('Word damaged picture data in a reusable text-box slot does not notify', async () => {
  const inside = '\x01\r',
    table = txbx(inside);
  table.writeUInt16LE(1, 16);
  const r = await convert(doc({ text: 'Body\r', stories: { textboxes: inside }, tables: [[56, table]] }));
  assert.equal(texts(r.xml), 'Body');
  assert.ok(!r.losses.some(x => x.notify));
});
test('Word picture transparency and unequal PICF borders survive in VML', async () => {
  const data = picture([[263, 0x332211]]);
  data.writeUInt32LE(8 | (1 << 8) | (6 << 16), 46);
  data.writeUInt32LE(16 | (3 << 8) | (2 << 16), 54);
  const r = await convert(inlineDoc(data));
  assert.match(r.xml, /chromakey="#112233"/);
  assert.match(r.xml, /o:bordertopcolor="#FF0000"/);
  assert.match(r.xml, /<w10:bordertop type="single" width="8"/);
  assert.match(r.xml, /<w10:borderbottom type="double" width="16"/);
  assert.equal(r.stats.pictures, 1);
});
test('Word WordArt keeps text and explicit fitting flags without inferring its dimensions', async () => {
  const text = '\x08Body\r',
    s = shape(1025, 136, [
      [192, Buffer.from('Watermark\0', 'utf16le')],
      [255, 0xffff4100],
      [195, 20 * 65536]
    ]),
    r = await convert(
      doc({
        text,
        tables: [
          [50, drawing([s])],
          [40, spa([{ id: 1025, cp: 0 }], text.length)]
        ]
      })
    );
  assert.match(r.xml, /<v:textpath on="t" fitshape="f" fitpath="t" trim="f" xscale="f" string="Watermark"/);
  assert.match(r.xml, /font-size:20pt/);
  assert.match(r.xml, /width:72pt;height:36pt/);
});
test('Word embedded objects keep their ObjectPool WMF preview and report the lost functionality', async () => {
  const text = '\x01\r',
    bytes = doc({
      text,
      chpx: [{ start: 2048, end: 2052, ops: [sprm(0x6a03, 1234), sprm(0x080a, 1), sprm(0x0856, 1)] }]
    }),
    cfb = B.cfb(bytes),
    pic = Buffer.alloc(100),
    meta = Buffer.alloc(32);
  pic.writeInt32LE(1440, 20);
  pic.writeInt32LE(720, 24);
  pic.writeInt32LE(1000, 44);
  pic.writeInt32LE(1000, 48);
  meta.writeUInt16LE(8, 0);
  meta.writeInt16LE(2540, 2);
  meta.writeInt16LE(1270, 4);
  meta.writeUInt16LE(1, 8);
  meta.writeUInt16LE(9, 10);
  meta.writeUInt16LE(0x300, 12);
  meta.writeUInt32LE(12, 14);
  meta.writeUInt32LE(3, 20);
  meta.writeUInt32LE(3, 26);
  const original = B.cfb;
  B.cfb = () => ({
    ...cfb,
    stream: name =>
      name === 'ObjectPool/_1234/\x03PIC' ? pic : name === 'ObjectPool/_1234/\x03META' ? meta : cfb.stream(name)
  });
  try {
    const r = await convert(bytes);
    assert.match(r.xml, /<wp:inline/);
    assert.match(r.xml, /<wp:extent cx="914400" cy="457200"/);
    assert.equal(r.stats.pictures, 1);
    assert.deepEqual(Buffer.from(await r.zip.get(mediaNames(r)[0]).bytes()), meta.subarray(8));
    assert.ok(r.losses.some(x => x.notify && x.what === 'An embedded object was kept only as a picture'));
  } finally {
    B.cfb = original;
  }
});
