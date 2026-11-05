import {test}from 'node:test';import assert from 'node:assert/strict';import fs from 'node:fs';import vm from 'node:vm';
import {load}from '../../tools/legacy/load.mjs';import {rec,container,u16,u32,scheme,slideAtom,presentation}from './ppt-fixtures.mjs';
const B=load(),P=B.ppt;
if(process.env.LEGACY_BEFORE_D2)vm.runInThisContext(fs.readFileSync(process.env.LEGACY_BEFORE_D2,'utf8'));
if(process.env.LEGACY_BEFORE_D2_TEXT)for(const file of ['ppt-text','ppt-text-design','ppt-text-write','ppt-write'])vm.runInThisContext(fs.readFileSync(process.env.LEGACY_BEFORE_D2_TEXT+'/'+file+'.js','utf8'));
if(process.env.LEGACY_BEFORE_D2_SDK)for(const file of ['ppt-text-write','ppt-theme'])vm.runInThisContext(fs.readFileSync(process.env.LEGACY_BEFORE_D2_SDK+'/sdk-'+file+'.js','utf8'));
const concat=(...parts)=>Buffer.concat(parts),chars=(s,type=4)=>[rec(3999,u32(type)),rec(4000,Buffer.from(s,'utf16le'))];
const shape=(id,records,{ph,client=[],anchor=[200,100,2100,1200],child=false,type=202,properties=[]}={})=>container(0xf004,[rec(0xf00a,concat(u32(id),u32(0xa00)),type,2),rec(child?0xf00f:0xf010,concat(...anchor.map(u32))),...(properties.length?[rec(0xf00b,concat(...properties.map(([id,v])=>concat(u16(id),u32(v)))),properties.length,3)]:[]),rec(0xf00d,concat(...records)),rec(0xf011,concat(...(ph?[rec(3011,concat(u32(ph[0]),Buffer.from([ph[1],0,0,0])))]:[]),...client))]);
const drawing=shapes=>rec(1036,container(0xf002,shapes));
const make=(shapes,{masterRecords=[],textRecords=[],environmentRecords=[]}={})=>presentation({environmentRecords,pages:{masters:[{id:0x80000000,ref:2,record:container(1016,[slideAtom(),scheme(P.defaultScheme),...masterRecords])}],slides:[{id:256,ref:3,textRecords,record:container(1006,[slideAtom(0x80000000),drawing(shapes)])}],notes:[]}});
async function convert(f){const result=await P.convert(f.bytes),z=await L.zip.read(result.pkg),xml=await z.get('ppt/slides/slide1.xml').text();return {...result,z,xml};}
const count=(xml,s)=>(xml.match(new RegExp(s,'g'))||[]).length;

test('PPT text follows shape order and outline indices without duplicate text',async()=>{
 const f=make([shape(2051,[rec(3998,u32(1))],{ph:[1,14]}),shape(2050,[rec(3998,u32(0))],{ph:[0,13]})],{textRecords:[...chars('Title',0),...chars('Body',1)]});
 const {xml,losses}=await convert(f);assert.equal(count(xml,'<a:t>Body</a:t>'),1);assert.equal(count(xml,'<a:t>Title</a:t>'),1);assert.ok(xml.indexOf('>Body<')<xml.indexOf('>Title<'));assert.match(xml,/<p:ph type="body" idx="1"/);assert.match(xml,/id="2051"/);assert.ok(!losses.some(l=>l.id.includes('text')&&l.notify));
});
test('PPT UTF-16 run lengths, soft breaks and paragraph styles survive conversion',async()=>{
 const text='A😀\vB\rSecond',style=rec(4001,concat(u32(6),u16(0),u32(1),u16(0),u32(7),u16(1),u32(0x581),u16(1),u16(42),u16(576),u16(288),u32(3),u32(0x20001),u16(1),u16(28),u32(10),u32(0x40001),u16(0),u32(0xfe332211)));
 const {xml}=await convert(make([shape(1026,[...chars(text),style])]));assert.match(xml,/<a:t>A😀<\/a:t>/);assert.equal(count(xml,'<a:br>'),1);assert.match(xml,/<a:pPr lvl="1" marL="914400" indent="-457200"/);assert.match(xml,/<a:buChar char="\*"/);assert.match(xml,/sz="2800" b="1"/);assert.match(xml,/val="112233"/);assert.equal(count(xml,'<a:p>'),2);
});
test('PPT TextBytesAtom contains low Unicode bytes, not Windows-1252',async()=>{
 const {xml}=await convert(make([shape(1026,[rec(3999,u32(4)),rec(4008,Buffer.from([65,128,233]))])]));assert.match(xml,/<a:t>A\u0080é<\/a:t>/);
});
test('PPT master levels inherit and TextRuler supplies margins and tabs',async()=>{
 const master=rec(4003,concat(u16(2),u32(0),u32(0x30000),u16(0),u16(30),u32(0),u32(1),u16(1)),1),style=rec(4001,concat(u32(5),u16(1),u32(0),u32(5),u32(0))),ruler=rec(4006,concat(u32(0x211),u16(720),u16(500),u16(200)));
 const {xml,z}=await convert(make([shape(2050,[...chars('Text',1),style,ruler],{ph:[1,14]})],{masterRecords:[master]}));assert.match(xml,/sz="3000" b="1"/);assert.match(xml,/marL="793750" indent="-476250"/);assert.match(xml,/defTabSz="1143000"/);assert.match(xml,/typeface="Arial"/);assert.match(await z.get('ppt/slideMasters/slideMaster1.xml').text(),/<p:bodyStyle><a:lvl1pPr/);
});
test('PPT ninth-era numbered bullets retain their type and start number',async()=>{
 const extension=container(5000,[container(5002,[rec(4026,Buffer.from('___PPT9','utf16le')),rec(5003,rec(4012,concat(u32(0x3000000),u16(1),u16(3),u16(4),u32(0),u32(0))))])]);
 const {xml}=await convert(make([shape(2050,chars('Fourth',1),{ph:[1,14],client:[extension]})]));assert.match(xml,/<a:buAutoNum type="arabicPeriod" startAt="4"/);
});
test('PPT fields remain editable fields with source slide numbering',async()=>{
 const {xml}=await convert(make([shape(2050,[...chars('Page *'),rec(4056,u32(5))])]));assert.match(xml,/<a:fld id="\{[0-9A-F-]{36}\}" type="slidenum">/);assert.match(xml,/<a:t>5<\/a:t>/);assert.equal(count(xml,'<a:fld '),1);
});
test('PPT broken formatting and an unresolved outline do not discard other text',async()=>{
 const {xml,losses}=await convert(make([shape(2050,[...chars('Still here'),rec(4001,concat(u32(0),u16(0),u32(0)))]),shape(2051,[rec(3998,u32(99))]),shape(2052,chars('Other box'))]));assert.match(xml,/>Still here</);assert.match(xml,/>Other box</);assert.ok(losses.some(l=>l.id==='legacy-presentation-format'&&l.notify));assert.ok(losses.some(l=>l.id==='legacy-presentation-reference'&&l.notify));
});
test('PPT grouped text keeps child coordinates and quarter-turn anchor semantics',async()=>{
 const group=container(0xf003,[container(0xf004,[rec(0xf00a,concat(u32(1025),u32(1)),0,2),rec(0xf009,concat(...[0,0,1000,1000].map(u32))),rec(0xf010,concat(...[0,0,2000,2000].map(u32)))]),shape(1026,chars('Turned'),{child:true,anchor:[0,0,1000,500],properties:[[4,90*65536]]})]);
 const {xml}=await convert(make([group]));assert.match(xml,/<p:grpSp>/);assert.match(xml,/<a:chExt cx="1000" cy="1000"/);assert.match(xml,/rot="5400000"><a:off x="250" y="-250"\/><a:ext cx="500" cy="1000"/);assert.match(xml,/>Turned</);
});
test('PPT trailing empty paragraphs and XML metacharacters are retained',async()=>{
 const {xml}=await convert(make([shape(2050,chars('<one> & two\r'))]));assert.equal(count(xml,'<a:p>'),2);assert.match(xml,/&lt;one&gt; &amp; two/);
});
test('PPT native horizontal and vertical text flow and font rotation are distinct',async()=>{
 const {xml}=await convert(make([shape(2050,chars('Horizontal'),{properties:[[136,4],[137,3]]}),shape(2051,chars('Vertical'),{properties:[[136,3]]})]));assert.match(xml,/vert="horz" rot="-5400000"/);assert.match(xml,/vert="vert"/);
});
test('PPT composite-master static text survives beside a retained layout',async()=>{
 const layout=B.xmlDeclaration+'<p:sldLayout xmlns:p="'+B.ns.p+'" xmlns:a="'+B.ns.a+'" type="title"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sldLayout>',zip=new Uint8Array(await(await L.zip.write([{name:'ppt/slideLayouts/slideLayout1.xml',data:layout}])).arrayBuffer());
 const f=presentation({pages:{masters:[{id:0x80000000,ref:2,record:container(1016,[slideAtom(),rec(1052,u32(1))])},{id:0x80000001,ref:4,record:container(1016,[slideAtom(),rec(1053,u32(1)),rec(1054,zip,1),drawing([shape(2050,chars('Added to this layout'))])])}],slides:[{id:256,ref:3,record:container(1006,[slideAtom(0x80000001),rec(1058,concat(u32(1),u32(1)))])}],notes:[]}});
 const {z}=await convert(f),xml=await z.get('ppt/slideLayouts/slideLayout1.xml').text();assert.equal(count(xml,'>Added to this layout<'),1);assert.match(xml,/<p:cNvPr id="2050"/);B.checkXml(xml);
});
test('PPT picture bullets retain image bytes and a relationship from the text part',async()=>{
 const image=await B.images.png(1,1,new Uint8Array([255,0,0,255])),blip=rec(0xf01e,concat(Buffer.alloc(16),Buffer.from([255]),Buffer.from(image)),0x6e0);
 const tag=entries=>container(5000,[container(5002,[rec(4026,Buffer.from('___PPT9','utf16le')),rec(5003,concat(...entries))])]);
 const master=rec(4003,concat(u16(1),u32(1),u16(1),u32(0)),1),master9=rec(4013,concat(u16(1),u32(0x800000),u16(0),u32(0)),1);
 const f=make([shape(2050,chars('Picture list',1),{ph:[1,14]})],{masterRecords:[master,tag([master9])]});
 // Rebuild with the document's tagged bullet collection, retaining the source pages.
 const withImage=presentation({pages:f.pages,documentRecords:[container(2000,[tag([container(2040,[rec(2041,concat(u16(6),blip),0)])])])]});
 const {z,xml,losses}=await convert(withImage);assert.match(xml,/<a:buBlip><a:blip r:embed="rId\d+"/);assert.deepEqual(await z.get('ppt/media/image1.png').bytes(),image);assert.match(await z.get('ppt/slides/_rels/slide1.xml.rels').text(),/media\/image1.png/);assert.ok(!losses.some(l=>l.id==='legacy-presentation-bullet'));
});
test('PPT later East Asian font references use their own font collection',async()=>{
 const tag=entries=>container(5000,[container(5002,[rec(4026,Buffer.from('___PPT10','utf16le')),rec(5003,concat(...entries))])]),face=Buffer.alloc(68);face.write('MS Mincho',0,'utf16le');face[64]=128;face[67]=192;
 const f=make([shape(2050,chars('日本語'),{client:[tag([rec(4017,concat(u32(0x1000000),u16(0)))])]})]);
 const extended=presentation({pages:f.pages,documentRecords:[container(2000,[tag([container(2006,[rec(4023,face,0)])])])]});const {xml}=await convert(extended);assert.match(xml,/<a:ea typeface="MS Mincho" pitchFamily="-64" charset="-128"/);assert.match(xml,/<a:latin typeface="Arial"/);
});
test('PPT date fields read the format directly after the text position',async()=>{
 const {xml,losses}=await convert(make([shape(2050,[...chars('* *'),rec(4087,concat(u32(0),Buffer.from([12,0,0,0]))),rec(4117,concat(u32(2),Buffer.from("'Fixed date'\0",'utf16le'),Buffer.alloc(102)))])]));
 assert.match(xml,/type="datetime13"/);assert.match(xml,/>Fixed date</);assert.ok(!losses.some(l=>l.id==='legacy-presentation-format'||l.id==='legacy-presentation-field'));assert.ok(!xml.includes('<a:t>*'));
});
test('PPT retained layout supplies the missing subtitle placeholder index',async()=>{
 const layout=B.xmlDeclaration+'<p:sldLayout xmlns:p="'+B.ns.p+'" xmlns:a="'+B.ns.a+'" type="title"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="3" name="Subtitle"/><p:cNvSpPr/><p:nvPr><p:ph type="subTitle" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp></p:spTree></p:cSld></p:sldLayout>',zip=new Uint8Array(await(await L.zip.write([{name:'ppt/slideLayouts/slideLayout1.xml',data:layout}])).arrayBuffer());
 const f=presentation({pages:{masters:[{id:0x80000000,ref:2,record:container(1016,[slideAtom(),rec(1052,u32(1)),rec(1054,zip,1)])}],slides:[{id:256,ref:3,record:container(1006,[slideAtom(0x80000000),rec(1058,concat(u32(1),u32(1))),drawing([shape(2050,chars('Subtitle',5),{ph:[-1,16],client:[rec(1055,u32(3))]})])])}],notes:[]}});
 const {xml}=await convert(f);assert.match(xml,/<p:ph type="subTitle" idx="1"/);
});
test('PPT modern header/footer placeholder records retain the placeholder role',async()=>{
 const {xml}=await convert(make([shape(2050,[...chars('*'),rec(4056,u32(0))],{client:[rec(1056,Buffer.from([8]))]})]));assert.match(xml,/<p:ph type="sldNum"/);assert.match(xml,/type="slidenum"/);
});
test('PPT embedded master text styles retain Office kerning and theme fonts',async()=>{
 const style='<p:txStyles xmlns:p="'+B.ns.p+'" xmlns:a="'+B.ns.a+'"><p:titleStyle><a:lvl1pPr><a:defRPr sz="4400" kern="1200"><a:latin typeface="+mj-lt"/></a:defRPr></a:lvl1pPr></p:titleStyle><p:bodyStyle/><p:otherStyle/></p:txStyles>',zip=new Uint8Array(await(await L.zip.write([{name:'drs/slideMasters/slideMaster1.xml',data:style}])).arrayBuffer());
 const {z,stats}=await convert(make([shape(2050,chars('Title',0))],{masterRecords:[rec(1059,zip)]}));const xml=await z.get('ppt/slideMasters/slideMaster1.xml').text();assert.match(xml,/kern="1200"/);assert.match(xml,/typeface="\+mj-lt"/);assert.equal(stats.storedTextStyles,1);B.checkXml(xml);
});
test('PPT symbol text keeps the symbol face when an ANSI fallback is present',async()=>{
 const symbol=Buffer.alloc(68);symbol.write('Symbol',0,'utf16le');symbol[64]=2;
 const style=rec(4001,concat(u32(2),u16(0),u32(0),u32(2),u32(0x410000),u16(1),u16(0))),f=make([shape(2050,[...chars('b'),style])]);
 const {xml}=await convert(presentation({pages:f.pages,fontRecords:[rec(4023,symbol,1)]}));assert.match(xml,/<a:latin typeface="Symbol"/);assert.match(xml,/>\uf062</);
});
test('PPT placeholders inherit individual master frame flags and text insets',async()=>{
 const parent=shape(1026,chars('Master',0),{type:1,ph:[0,1],properties:[[447,0x110001],[511,0x90001],[129,182880]]});
 const {xml}=await convert(make([shape(2050,chars('Title',0),{type:1,ph:[0,13],properties:[[769,1026],[447,0x10000],[511,0x10000]]})],{masterRecords:[drawing([parent])]}));
 assert.match(xml,/<a:avLst\/><\/a:prstGeom><a:noFill\/><a:ln><a:noFill\/><\/a:ln>/);assert.match(xml,/<a:bodyPr lIns="182880"/);
});
test('PPT groups keep large native child coordinates within DrawingML limits',async()=>{
 const group=container(0xf003,[container(0xf004,[rec(0xf00a,concat(u32(1025),u32(1)),0,2),rec(0xf009,concat(...[0,0,4572000,4572000].map(u32))),rec(0xf010,concat(...[0,0,2000,2000].map(u32)))]),shape(1026,chars('Grouped'),{child:true,anchor:[0,0,3657600,3657600]})]);
 const {xml}=await convert(make([group]));assert.match(xml,/<a:chExt cx="4572000" cy="4572000"/);B.checkXml(xml,node=>{if(node.name==='a:ext'||node.name==='a:chExt')for(const key of ['cx','cy'])assert.ok(+node.attrs[key]<=2147483647);});
});
test('PPT unmaterialized header/footer content is an explicit loss until D3',async()=>{
 const f=make([shape(2050,chars('Body'))]),headers=container(4057,[rec(4058,concat(u16(0),u16(32))),rec(4026,Buffer.from('Confidential','utf16le'),2)],3);
 const {losses}=await convert(presentation({pages:f.pages,documentRecords:[headers]}));assert.ok(losses.some(l=>l.id==='legacy-slide-headers-footers'&&l.notify&&l.where==='Slide 1'));
});
test('PPT right-parenthesis numbering uses the DrawingML scheme name',async()=>{
 const extension=container(5000,[container(5002,[rec(4026,Buffer.from('___PPT9','utf16le')),rec(5003,rec(4012,concat(u32(0x3000000),u16(1),u16(9),u16(1),u32(0),u32(0))))])]);
 const {xml}=await convert(make([shape(2050,chars('First',1),{ph:[1,14],client:[extension]})]));assert.match(xml,/<a:buAutoNum type="alphaLcParenR" startAt="1"/);
});
