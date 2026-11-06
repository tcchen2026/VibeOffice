import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import { inflateSync, deflateSync } from 'node:zlib';
import { load } from '../../tools/legacy/load.mjs';
const B = load(), A = B.officeart;
if (process.env.IMAGES_SOURCE) vm.runInThisContext(fs.readFileSync(process.env.IMAGES_SOURCE, 'utf8'));
if (process.env.SHAPE_PATH_SOURCE) vm.runInThisContext(fs.readFileSync(process.env.SHAPE_PATH_SOURCE, 'utf8'));
if (process.env.SHAPES_SOURCE) vm.runInThisContext(fs.readFileSync(process.env.SHAPES_SOURCE, 'utf8'));
const record = (type, data, instance = 0, version = 0) => { const h = Buffer.alloc(8); h.writeUInt16LE(instance * 16 + version); h.writeUInt16LE(type, 2); h.writeUInt32LE(data.length, 4); return Buffer.concat([h, data]); };
const pngPixels = bytes => {
  const b = Buffer.from(bytes); let p = 8, compressed = []; assert.equal(b.subarray(0, 8).toString('hex'), '89504e470d0a1a0a');
  while (p < b.length) { const n = b.readUInt32BE(p), kind = b.subarray(p + 4, p + 8).toString(); assert.equal(globalThis.L.zip.crc32(b.subarray(p + 4, p + 8 + n)), b.readUInt32BE(p + 8 + n)); if (kind === 'IDAT') compressed.push(b.subarray(p + 8, p + 8 + n)); p += n + 12; }
  return inflateSync(Buffer.concat(compressed));
};
function dib(width, height, bits, data, colours = [], compression = 0) {
  const h = Buffer.alloc(40); h.writeUInt32LE(40); h.writeInt32LE(width, 4); h.writeInt32LE(height, 8); h.writeUInt16LE(1, 12); h.writeUInt16LE(bits, 14); h.writeUInt32LE(compression, 16); h.writeUInt32LE(colours.length, 32);
  return Buffer.concat([h, ...colours.map(([r,g,b]) => Buffer.from([b,g,r,0])), data]);
}
test('DIB rows and indexed pixels are converted to independently decodable PNG bytes', async () => {
  const bytes = dib(2, 2, 24, Buffer.from([255,0,0,255,255,255,0,0,0,0,255,0,255,0,0,0]));
  const p = await B.images.dib(bytes); assert.equal(p.ext, 'png');
  assert.deepEqual([...pngPixels(p.bytes)], [0,255,0,0,255,0,255,0,255,0,0,0,255,255,255,255,255,255]);
  const indexed = await B.images.dib(dib(2, -1, 1, Buffer.from([0x80,0,0,0]), [[255,0,0],[0,0,255]]));
  assert.deepEqual([...pngPixels(indexed.bytes)], [0,0,0,255,255,255,0,0,255]);
});
test('RLE bitmap pixels and deltas remain bounded', async () => {
  const bytes = dib(3, 1, 8, Buffer.from([0,3,1,0,1,0,0,1]), [[0,0,0],[255,0,0]], 1);
  assert.deepEqual([...pngPixels((await B.images.dib(bytes)).bytes)], [0,255,0,0,255,0,0,0,255,255,0,0,255]);
  const bad = dib(3, 1, 8, Buffer.from([4,1,0,1]), [[0,0,0],[255,0,0]], 1);
  await assert.rejects(() => B.images.dib(bad), { code: 'DIB_PIXELS' });
  await assert.rejects(() => B.images.dib(dib(1000000,1000000,24,Buffer.alloc(0))), { code: 'IMAGE_LIMIT' });
});
test('PNG uses compressed DEFLATE and falls back to identical stored pixels', async () => {
  const rgba = new Uint8Array(128 * 128 * 4).fill(255), compressed = await B.images.png(128, 128, rgba);
  assert.ok(compressed.length < 1000, 'flat picture must compress, got ' + compressed.length + ' bytes');
  const native = globalThis.CompressionStream;
  try {
    globalThis.CompressionStream = undefined;
    const fallback = await B.images.png(128, 128, rgba);
    assert.ok(fallback.length > rgba.length);
    assert.deepEqual(pngPixels(fallback), pngPixels(compressed));
  } finally { globalThis.CompressionStream = native; }
});
test('BLIPs preserve picture bytes and bound compressed metafiles', async () => {
  const png = await B.images.png(1, 1, new Uint8Array([20,30,40,255]));
  const image = await A.blip(A.record(record(0xf01e, Buffer.concat([Buffer.alloc(32), Buffer.from([255]), png]), 0x6e1)));
  assert.deepEqual([...image.bytes], [...png]); const emf = Buffer.from('metafile-bytes'), compressed = deflateSync(emf), h = Buffer.alloc(50);
  h.writeUInt32LE(emf.length, 16); h.writeUInt32LE(compressed.length, 44); h[49] = 254;
  const media = await A.blip(A.record(record(0xf01a, Buffer.concat([h,compressed]), 0x3d4)));
  assert.equal(media.ext, 'emf'); assert.deepEqual(Buffer.from(media.bytes), emf);
  await assert.rejects(() => B.images.inflate(deflateSync(Buffer.alloc(100000)), 20), { code: 'IMAGE_LIMIT' });
  await assert.rejects(() => B.images.inflate(Buffer.from([0,0]), 10), { code: 'IMAGE_COMPRESSION' });
});
test('OfficeArt nesting and complex-array lengths are bounded', () => {
  let bytes = record(0xf00a, Buffer.alloc(8)); for (let i=0;i<130;i++) bytes = record(0xf003,bytes,0,15);
  assert.throws(() => A.records(bytes), { code: 'OFFICEART_DEPTH' });
  const p = Buffer.alloc(16); p.writeUInt16LE(0x8000|325); p.writeUInt32LE(4,2); p.writeUInt16LE(1,6); p.writeUInt16LE(1,8); p.writeInt16LE(-16,10); p.writeInt16LE(10,12); p.writeInt16LE(20,14);
  const props = A.properties(A.record(record(0xf00b,p,1,3))), vertex = A.array(props.get(325))[0];
  assert.equal(vertex.i16(0),10); assert.equal(vertex.i16(2),20);
  p.writeUInt32LE(999,2); assert.throws(() => A.properties(A.record(record(0xf00b,p,1,3))), { code:'TRUNCATED' });
});
test('bad unused BLIPs are lazy and do not hide sound pictures', async () => {
  const data=Buffer.alloc(36); data.writeUInt32LE(8,20); data.writeUInt32LE(1,24); data.writeUInt32LE(9999,28);
  const entries=A.store(A.records(record(0xf001,record(0xf007,data),1,15)),new Uint8Array(0));
  assert.equal(entries.length,1); await assert.rejects(entries[0].read(),{code:'TRUNCATED'});
});
test('host anchors, tertiary table properties, groups and shape transforms keep their coordinate systems', () => {
  const group=Buffer.alloc(8); group.writeUInt32LE(1024); group.writeUInt32LE(1|64,4);
  const anchor=Buffer.alloc(8); [10,20,100,60].forEach((v,i)=>anchor.writeInt16LE(v,i*2));
  const prop=Buffer.alloc(6);prop.writeUInt16LE(927);prop.writeUInt32LE(1,2);
  const tree=A.records(record(0xf003,record(0xf004,Buffer.concat([record(0xf00a,group,1,2),record(0xf010,anchor),record(0xf122,prop,1,3)]),0,15),0,15));
  const s=A.shapes(tree,'ppt')[0];assert.equal(s.table,true);assert.deepEqual(s.anchor,{kind:'rect',top:10,left:20,right:100,bottom:60,units:576});
  assert.match(B.shapes.transform(s,{x:1,y:2,cx:3,cy:4},true),/flipH="1"/);assert.match(B.shapes.geometry(s),/prst="rect"/);
  s.properties.set(327,{value:100});assert.match(B.shapes.geometry(s),/<a:custGeom>/);
});
test('native guides preserve an adjusted arrow outline without an invalid partial preset list', () => {
  const shape={type:13,properties:new Map([[327,{value:14000}],[328,{value:6000}]]),flags:0};
  const path=B.shapePath(shape), xml=B.shapes.geometry(shape);
  assert.ok(path.paths.length);assert.match(xml,/<a:custGeom>/);assert.doesNotMatch(xml,/NaN|Infinity|undefined/);
  assert.ok(path.paths.flatMap(p=>p.commands.flatMap(c=>c.pts)).every(p=>p.every(Number.isFinite)));
  assert.match(B.shapes.vml(shape,{x:0,y:0,cx:100000,cy:100000}),/<v:shape/);
});
test('all native definitions tolerate unused trailing commands without inventing points', () => {
  for (const type of Object.keys(B.nativeGeometry).map(Number)) {
    const path = B.shapePath({ type, properties: new Map(), flags: 0 });
    assert.ok(path.paths.length, 'native shape ' + type);
    assert.ok(path.paths.flatMap(p => p.commands.flatMap(c => c.pts)).every(p => p.every(Number.isFinite)), 'native shape ' + type);
  }
  const segments = Buffer.alloc(10); segments.writeUInt16LE(2); segments.writeUInt16LE(2, 2); segments.writeUInt16LE(2, 4);
  segments.writeUInt16LE(0x4000, 6); segments.writeUInt16LE(10, 8);
  assert.throws(() => B.shapePath({type:1,properties:new Map([[326,{data:segments}]])}), {code:'SHAPE_PATH'});
});
test('custom guides reject cycles and resolve signed coordinates and ellipse arcs', () => {
  const arr=(rows,size)=>{const b=Buffer.alloc(6+rows.length*size);b.writeUInt16LE(rows.length);b.writeUInt16LE(rows.length,2);b.writeUInt16LE(size,4);rows.forEach((row,i)=>row.forEach((v,j)=>size===8&&row.length===2?b.writeUInt32LE(v>>>0,6+i*size+j*4):b.writeUInt16LE(v,6+i*size+j*2)));return{data:b};};
  const props=new Map([[325,arr([[0,0],[100,100],[100,50],[50,100]],8)],[326,arr([[0xa504],[0x8000]],2)],[322,{value:100}],[323,{value:100}]]);
  const path=B.shapePath({type:0,properties:props});assert.ok(path.paths[0].commands.some(c=>c.op==='C'));assert.deepEqual(path.paths[0].commands[0].pts[0],[100,50]);
  props.set(325,arr([[0x80000000,0]],8));props.set(326,arr([[0x4000],[0x8000]],2));props.set(342,arr([[0x2000,0x400,0,0]],8));
  assert.throws(()=>B.shapePath({type:0,properties:props}),{code:'SHAPE_GUIDES'});
});
test('custom connection points and text bounds scale with the shape when resized', () => {
  const array=(values,size)=>{const data=Buffer.alloc(6+values.length*size);data.writeUInt16LE(values.length);data.writeUInt16LE(values.length,2);data.writeUInt16LE(size,4);values.forEach((v,i)=>v.forEach((n,j)=>data.writeInt32LE(n,6+i*size+j*4)));return{data};};
  const s={type:5,flags:0,properties:new Map([[337,array([[5400,10800],[21600,5400]],8)],[338,array([[180*65536],[90*65536]],4)]])};
  assert.deepEqual(B.shapePath(s).connections,[{x:5400,y:10800,angle:180},{x:21600,y:5400,angle:90}]);
  const xml=B.shapes.geometry(s),guides=new Map();
  B.checkXml('<root xmlns:a="'+B.ns.a+'">'+xml+'</root>',node=>{if(node.name==='a:gd')guides.set(node.attrs.name,node.attrs.fmla);});
  const resolve=(name,w,h)=>{const[op,axis,n,d]=guides.get(name).split(' ');assert.equal(op,'*/');return (axis==='w'?w:h)*Number(n)/Number(d);};
  assert.match(xml,/<a:cxn ang="10800000"><a:pos x="cx0" y="cy0"/);
  assert.equal(resolve('cx0',4000000,2000000),1000000);assert.equal(resolve('cy1',4000000,2000000),500000);
  assert.equal(resolve('txr',8000000,2160000),6000000);assert.equal(resolve('txb',8000000,2160000),1800000);
});
