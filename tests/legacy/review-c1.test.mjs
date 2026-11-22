import {test} from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import {load} from '../../tools/legacy/load.mjs';
const B=load();
if(process.env.LEGACY_REVIEW_BEFORE)for(const name of ['shapes','package','property-parts'])vm.runInThisContext(fs.readFileSync(process.env.LEGACY_REVIEW_BEFORE+'/'+name+'.js','utf8'));
test('rotated Office anchors are unswapped about their centre in DML and VML',()=>{
  const shape={id:1,type:1,flags:0,properties:new Map([[4,{value:90*65536}]])},rect={x:100,y:200,cx:1000,cy:400};
  assert.match(B.shapes.transform(shape,rect),/<a:off x="400" y="-100"\/><a:ext cx="400" cy="1000"\/>/);
  const vml=B.shapes.vml(shape,rect);assert.ok(vml.includes('width:'+400/12700+'pt;height:'+1000/12700+'pt'));
  assert.deepEqual(rect,{x:100,y:200,cx:1000,cy:400});
  shape.properties.set(4,{value:45*65536});assert.equal(B.shapes.anchorRect(shape,rect).cx,400);
  shape.properties.set(4,{value:135*65536});assert.equal(B.shapes.anchorRect(shape,rect).cx,1000);
});
test('Office gradient kinds 4 and 7 are linear, 5 and 6 follow bounds and outline',()=>{
  for(const type of [4,5,6,7]){
    const shape={id:1,type:1,flags:0,properties:new Map([[384,{value:type}],[395,{value:-30*65536}]])};
    const dml=B.shapes.style(shape),vml=B.shapes.vml(shape,{x:0,y:0,cx:1000,cy:1000});
    if(type===4||type===7){assert.match(dml,/<a:lin ang="7200000"/);assert.match(vml,/type="gradient"/);}
    else{assert.match(dml,new RegExp('<a:path path="'+(type===5?'rect':'shape')+'"'));assert.match(vml,/type="gradientRadial"/);}
  }
});
test('linear gradient angle, focus and rotation match Office-written embedded shape XML',()=>{
  // POI bug62591.ppt, shape 1032; its metro XML has angle 5400000,
  // white at 0 and FF931D at 100000, rotWithShape=0, scaled=1.
  const shape={id:1032,type:1,flags:0,properties:new Map([[384,{value:7}],[385,{value:0xffffff}],[387,{value:0x1d93ff}],[396,{value:100}],[447,{value:0x200000}]])};
  const xml=B.shapes.style(shape);assert.match(xml,/<a:gradFill rotWithShape="0">/);assert.match(xml,/<a:gs pos="0"><a:srgbClr val="FFFFFF"/);assert.match(xml,/<a:gs pos="100000"><a:srgbClr val="FF931D"/);assert.match(xml,/<a:lin ang="5400000" scaled="1"/);
  assert.match(B.shapes.vml(shape,{x:0,y:0,cx:1000,cy:500}),/focus="0%"/);shape.properties.set(396,{value:0});assert.match(B.shapes.vml(shape,{x:0,y:0,cx:1000,cy:500}),/focus="100%"/);
  shape.properties.set(396,{value:50});const focused=B.shapes.style(shape);assert.match(focused,/<a:gs pos="50000">/);
});
test('media content types are extension Defaults shared by every matching part',async()=>{
  const p=new B.Package();p.add('word/media/image1.png',new Uint8Array([1]),'image/png');p.add('word/media/image2.png',new Uint8Array([2]),'image/png');p.add('word/document.xml','<x/>','application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml');
  const zip=await L.zip.read(await p.finish()),xml=await zip.get('[Content_Types].xml').text();
  assert.equal((xml.match(/Extension="png"/g)||[]).length,1);assert.doesNotMatch(xml,/PartName="[^\"]+\.png"/);assert.match(xml,/PartName="\/word\/document.xml"/);
});
test('core property dates use Office second precision',()=>{
  const b=Buffer.alloc(76);b.writeUInt16LE(0xfffe);b.writeUInt32LE(1,24);b.writeUInt32LE(48,44);b.writeUInt32LE(28,48);b.writeUInt32LE(1,52);b.writeUInt32LE(12,56);b.writeUInt32LE(16,60);b.writeUInt32LE(64,64);b.writeBigUInt64LE((BigInt(Date.parse('2020-01-01T12:00:00.123Z'))+11644473600000n)*10000n,68);
  const result=B.propertyParts({stream:name=>name==='\x05SummaryInformation'?b:null});assert.match(result.parts[0].data,/>2020-01-01T12:00:00Z</);assert.doesNotMatch(result.parts[0].data,/\.123Z/);
});
