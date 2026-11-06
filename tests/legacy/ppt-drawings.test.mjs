import {test} from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import {load} from '../../tools/legacy/load.mjs';
import {rec,container,u16,u32,scheme,slideAtom,presentation,cfb} from './ppt-fixtures.mjs';
const B=load(),P=B.ppt;
if(process.env.LEGACY_BEFORE_D3)for(const name of ['officeart','shape-path','shapes','ppt-read','ppt-structure','ppt-theme','ppt-text-read','ppt-text','ppt-text-design','ppt-text-write','ppt-write'])vm.runInThisContext(fs.readFileSync(process.env.LEGACY_BEFORE_D3+'/'+name+'.js','utf8'));
const concat=(...a)=>Buffer.concat(a),chars=s=>[rec(3999,u32(4)),rec(4000,Buffer.from(s,'utf16le'))];
const prop=items=>rec(0xf00b,concat(...items.map(([id,value])=>concat(u16(id),u32(value)))),items.length,3);
const shape=(id,type=1,options={})=>container(0xf004,[rec(0xf00a,concat(u32(id),u32(options.flags??0xa00)),type,2),rec(options.child?0xf00f:0xf010,concat(...(options.anchor||[100,200,1200,600]).map(u32))),prop(options.properties||[]),...(options.text==null?[]:[rec(0xf00d,concat(...chars(options.text),...(options.textRecords||[])))]),...(options.client?[rec(0xf011,concat(...options.client))]:[])]);
const make=(shapes,options={})=>presentation({documentRecords:options.documentRecords||[],pages:{masters:options.masters||[{id:0x80000000,ref:2,record:container(1016,[slideAtom(),scheme(P.defaultScheme),...(options.masterRecords||[])])}],slides:[{id:256,ref:3,record:container(1006,[slideAtom(0x80000000,0,options.flags??7,16),rec(1036,container(0xf002,[...shapes,...(options.drawingRecords||[])])),...(options.slideRecords||[])])}],notes:[]}});
const blip=bytes=>rec(0xf01e,concat(Buffer.alloc(16),Buffer.from([255]),bytes),0x6e0);
const bse=bytes=>{const h=Buffer.alloc(36);h[0]=6;h[1]=6;h.writeUInt32LE(bytes.length,20);h.writeUInt32LE(1,24);return rec(0xf007,concat(h,bytes),6,2);};
const store=entries=>rec(1035,container(0xf000,[container(0xf001,entries,entries.length)]));
async function convert(f){const result=await P.convert(f.bytes),zip=await L.zip.read(result.pkg);return {...result,zip,xml:await zip.get('ppt/slides/slide1.xml').text()};}

test('PPT drawing-only geometry, colors and line arrows survive without text',async()=>{
 const r=await convert(make([shape(2050,3,{properties:[[385,0x332211],[448,0x665544],[459,25400],[465,1]]})]));
 assert.match(r.xml,/<p:sp>/);assert.match(r.xml,/<a:prstGeom prst="ellipse"/);assert.match(r.xml,/val="112233"/);assert.match(r.xml,/val="445566"/);assert.match(r.xml,/<a:tailEnd type="triangle"/);assert.ok(!r.losses.some(l=>l.id==='legacy-slide-drawings'));
});
test('PPT native picture bytes, crop and flipped anchors become picture parts',async()=>{
 const image=await B.images.png(2,1,Uint8Array.from([255,0,0,255,0,0,255,255]));
 const r=await convert(make([shape(2050,75,{flags:0xa40,properties:[[260,1],[256,16384],[258,32768]]})],{documentRecords:[store([bse(blip(image))])]}));
 assert.match(r.xml,/<p:pic>/);assert.match(r.xml,/<a:srcRect l="50000" t="25000"/);assert.match(r.xml,/flipH="1"/);
 const name=[...r.zip.keys()].find(n=>/^ppt\/media\/.*\.png$/.test(n));assert.ok(name);assert.deepEqual(await r.zip.get(name).bytes(),image);assert.match(await r.zip.get('ppt/slides/_rels/slide1.xml.rels').text(),/relationships\/image/);
});
test('PPT a malformed picture-store entry loses only its own picture',async()=>{
 const image=await B.images.png(1,1,Uint8Array.from([1,2,3,255]));
 const r=await convert(make([shape(2050,75,{properties:[[260,1]]}),shape(2051,75,{properties:[[260,2]]})],{documentRecords:[store([rec(0xf007,Buffer.alloc(10),6,2),bse(blip(image))])]}));
 assert.equal((r.xml.match(/<p:pic>/g)||[]).length,1);assert.match(r.xml,/id="2051"/);assert.ok(r.losses.some(l=>l.id==='legacy-presentation-picture'&&l.notify));
});
test('PPT unsupported picture color adjustments are reported while the picture stays',async()=>{
 const image=await B.images.png(1,1,Uint8Array.from([20,30,40,255]));
 const r=await convert(make([shape(2050,75,{properties:[[260,1],[265,22938]]}),shape(2051,75,{properties:[[260,1],[264,65536]]})],{documentRecords:[store([bse(blip(image))])]}));
 assert.equal((r.xml.match(/<p:pic>/g)||[]).length,2);const loss=r.losses.find(l=>l.id==='legacy-presentation-picture-adjustments');assert.ok(loss?.notify);assert.deepEqual(loss.detail.map(d=>d.shape),[2050]);
});
test('PPT group coordinates and quarter-turn shape bounds retain non-text children',async()=>{
 const group=container(0xf003,[container(0xf004,[rec(0xf00a,concat(u32(1025),u32(1)),0,2),rec(0xf009,concat(...[0,0,1000,1000].map(u32))),rec(0xf010,concat(...[0,0,2000,2000].map(u32)))]),shape(1026,1,{child:true,anchor:[0,0,1000,500],properties:[[4,90*65536],[385,0xff]]})]);
 const r=await convert(make([group]));assert.match(r.xml,/<p:grpSp>/);assert.match(r.xml,/<a:chExt cx="1000" cy="1000"/);assert.match(r.xml,/rot="5400000"><a:off x="250" y="-250"\/><a:ext cx="500" cy="1000"/);assert.match(r.xml,/val="FF0000"/);
});
test('PPT connector solver endpoints retain forward shape identities',async()=>{
 const rule=rec(0xf012,concat(...[1,2051,2052,2050,0,2].map(u32)),0,1);
 const r=await convert(make([shape(2050,32),shape(2051),shape(2052)],{drawingRecords:[container(0xf005,[rule])]}));
 assert.match(r.xml,/<p:cxnSp>/);assert.match(r.xml,/<a:stCxn id="2051" idx="0"\/>/);assert.match(r.xml,/<a:endCxn id="2052" idx="2"\/>/);
});
test('PPT clockwise rectangle sites map to the same sides in DrawingML',async()=>{
 const rule=rec(0xf012,concat(...[1,2051,2052,2050,1,3].map(u32)),0,1);
 const r=await convert(make([shape(2050,32),shape(2051),shape(2052)],{drawingRecords:[container(0xf005,[rule])]}));
 assert.match(r.xml,/<a:stCxn id="2051" idx="3"\/>/);assert.match(r.xml,/<a:endCxn id="2052" idx="1"\/>/);
});
test('PPT connector bends use complete preset adjustments and ignore unused binary slots',async()=>{
 const r=await convert(make([shape(2050,34,{properties:[[327,5400],[328,75600],[329,0xffff0000]]}),shape(2051,35,{properties:[[328,16200]]}),shape(2052,32,{properties:[[327,0xffffffff]]})]));
 assert.equal((r.xml.match(/<p:cxnSp>/g)||[]).length,3);assert.doesNotMatch(r.xml,/<a:custGeom>/);
 assert.match(r.xml,/<a:prstGeom prst="bentConnector3"><a:avLst><a:gd name="adj1" fmla="val 25000"\/><\/a:avLst>/);
 assert.match(r.xml,/<a:prstGeom prst="bentConnector4"><a:avLst><a:gd name="adj1" fmla="val 50000"\/><a:gd name="adj2" fmla="val 75000"\/><\/a:avLst>/);
 assert.match(r.xml,/<a:prstGeom prst="straightConnector1"><a:avLst\/><\/a:prstGeom>/);
});
test('PPT explicit solid and gradient backgrounds survive instead of the scheme default',async()=>{
 const background=container(0xf004,[rec(0xf00a,concat(u32(1),u32(0x400)),1,2),prop([[385,0x113355],[386,32768]])]);
 const r=await convert(make([background],{flags:0}));assert.match(r.xml,/<p:bgPr><a:solidFill><a:srgbClr val="553311"><a:alpha val="50000"/);assert.ok(!r.losses.some(l=>l.id==='legacy-slide-background'));
});

const action=(id,{kind=0,action=4,jump=0,type=8}={})=>container(4082,[rec(4083,concat(u32(0),u32(id),Buffer.from([action,0,jump,0,type,0,0,0])))],kind);
const range=(a,b)=>rec(4063,concat(u32(a),u32(b)));
const link=(id,target,location)=>container(4055,[rec(4051,u32(id)),...(target==null?[]:[rec(4026,Buffer.from(target,'utf16le'),1)]),...(location==null?[]:[rec(4026,Buffer.from(location,'utf16le'),3)])]);
const group=(children,props=[[927,1]],anchor=[0,0,2000,1000])=>container(0xf003,[container(0xf004,[rec(0xf00a,concat(u32(1025),u32(1)),0,2),rec(0xf009,concat(...[0,0,2000,1000].map(u32))),rec(0xf010,concat(...anchor.map(u32))),prop(props)]),...children]);

test('PPT merged grouped cells become an editable table with text, fill and borders',async()=>{
 const table=group([shape(1026,1,{child:true,anchor:[0,0,1000,1000],text:'Merged',properties:[[385,0x00ff00]]}),shape(1027,1,{child:true,anchor:[1000,0,2000,500],text:'Top'}),shape(1028,1,{child:true,anchor:[1000,500,2000,1000],text:'Bottom'}),shape(1029,20,{child:true,anchor:[0,0,2000,0],properties:[[448,0xff],[459,25400]]})]);
 const r=await convert(make([table]));assert.match(r.xml,/<a:graphicData uri="http:\/\/schemas.openxmlformats.org\/drawingml\/2006\/table">/);assert.match(r.xml,/<a:tbl>/);assert.match(r.xml,/<a:tc rowSpan="2">/);assert.match(r.xml,/<a:tc vMerge="1">/);assert.equal((r.xml.match(/<a:t>Merged<\/a:t>/g)||[]).length,1);assert.match(r.xml,/<a:lnT w="25400"[^>]*><a:solidFill><a:srgbClr val="FF0000"/);assert.match(r.xml,/val="00FF00"/);assert.equal(r.stats.drawings.tables,1);assert.ok(!r.losses.some(l=>l.id==='legacy-presentation-table'));
});
test('PPT overlapping table cells retain their separate text and report the lost table',async()=>{
 const table=group([shape(1026,1,{child:true,anchor:[0,0,1500,1000],text:'One'}),shape(1027,1,{child:true,anchor:[1000,0,2000,1000],text:'Two'})]);
 const r=await convert(make([table]));assert.doesNotMatch(r.xml,/<a:tbl>/);assert.match(r.xml,/<a:t>One<\/a:t>/);assert.match(r.xml,/<a:t>Two<\/a:t>/);assert.ok(r.losses.some(l=>l.id==='legacy-presentation-table'&&l.notify));
});
test('PPT text hyperlinks split only their UTF-16 range and keep ScreenTips',async()=>{
 const tip=container(5000,[container(5002,[rec(4026,Buffer.from('___PPT9','utf16le')),rec(5003,container(4068,[rec(4051,u32(7)),rec(4026,Buffer.from('A useful tip','utf16le'))]))])]);
 const r=await convert(make([shape(2050,202,{text:'A😀 BC',textRecords:[action(7),range(1,3)]})],{documentRecords:[container(1033,[link(7,'https://example.invalid/a b%no')]),tip]}));
 assert.match(r.xml,/<a:hlinkClick tooltip="A useful tip"[^>]*\/><\/a:rPr><a:t>😀<\/a:t>/);assert.equal((r.xml.match(/<a:hlinkClick/g)||[]).length,1);assert.match(await r.zip.get('ppt/slides/_rels/slide1.xml.rels').text(),/Target="https:\/\/example.invalid\/a%20b%25no" TargetMode="External"/);assert.equal(r.stats.drawings.hyperlinks,1);
});
test('PPT source slide IDs and mouse-over jump actions stay distinct',async()=>{
 const r=await convert(make([shape(2050,1,{client:[action(3,{type:7}),action(0,{kind:1,action:3,jump:6})]})],{documentRecords:[container(1033,[link(3,null,'256,1,First')])]}));
 assert.match(r.xml,/<a:hlinkClick action="ppaction:\/\/hlinksldjump" r:id="rId\d+"\/>/);assert.match(r.xml,/<a:hlinkMouseOver action="ppaction:\/\/hlinkshowjump\?jump=endshow" r:id=""\/>/);assert.match(await r.zip.get('ppt/slides/_rels/slide1.xml.rels').text(),/relationships\/slide" Target="slide1.xml"/);
});
test('PPT an invalid text link loses only that link, keeping the text and other link',async()=>{
 const r=await convert(make([shape(2050,202,{text:'Bad Good',textRecords:[action(1),range(0,3),action(2),range(4,8)]})],{documentRecords:[container(1033,[link(1,'javascript:alert(1)'),link(2,'mailto:person@example.invalid')])]}));
 assert.match(r.xml,/<a:t>Bad <\/a:t>/);assert.match(r.xml,/<a:hlinkClick[^>]+\/><\/a:rPr><a:t>Good<\/a:t>/);assert.ok(r.losses.some(l=>l.id==='legacy-presentation-link'&&l.notify));assert.doesNotMatch(await r.zip.get('ppt/slides/_rels/slide1.xml.rels').text(),/javascript/);
});
test('PPT hyperlinks stored in the outline text apply to its referencing shape',async()=>{
 const f=make([container(0xf004,[rec(0xf00a,concat(u32(2050),u32(0xa00)),202,2),rec(0xf010,concat(...[0,0,2000,1000].map(u32))),rec(0xf00d,rec(3998,u32(0)))])]);
 const outline=[...chars('Outline link'),action(2),range(8,12)];f.pages.slides[0].textRecords=outline;
 const r=await convert(presentation({pages:f.pages,documentRecords:[container(1033,[link(2,'https://example.invalid')])]}));assert.match(r.xml,/<a:t>Outline <\/a:t>/);assert.match(r.xml,/<a:hlinkClick[^>]+\/><\/a:rPr><a:t>link<\/a:t>/);assert.equal(r.stats.drawings.hyperlinks,1);
});
test('PPT property-set hyperlink targets follow ExHyperlink order and survive a damaged later entry',async()=>{
 const typed=n=>concat(u32(3),u32(n)),string=s=>{const data=Buffer.from(s+'\0','utf16le');return concat(u32(31),u32(s.length+1),data,Buffer.alloc((4-data.length%4)%4));};
 const blob=concat(u32(18),typed(91),typed(92),typed(93),typed(7),string('https://example.invalid/first'),string(''),typed(81),typed(82),typed(83),typed(7),string(''),string('256,1,First'),u32(0xffffffff));
 const name=Buffer.from('_PID_HLINKS\0'),dictionary=concat(u32(1),u32(2),u32(name.length),name),cp=concat(u32(2),u16(1252),u16(0)),value=concat(u32(65),u32(blob.length),blob);
 const section=concat(u32(32+dictionary.length+cp.length+value.length),u32(3),u32(0),u32(32),u32(1),u32(32+dictionary.length),u32(2),u32(32+dictionary.length+cp.length),dictionary,cp,value),header=Buffer.alloc(48);header.writeUInt16LE(0xfffe);header.writeUInt32LE(1,24);header.set(Buffer.from('05d5cdd59c2e1b10939708002b2cf9ae','hex'),28);header.writeUInt32LE(48,44);
 const f=make([shape(2050,202,{text:'Web Slide Kept',textRecords:[action(7),range(0,3),action(99,{type:7}),range(4,9),action(30),range(10,14)]})],{documentRecords:[container(1033,[link(7),link(99),link(30,'mailto:person@example.invalid')])]});
 const r=await convert({bytes:cfb({'PowerPoint Document':f.stream,'Current User':f.user,'\x05DocumentSummaryInformation':concat(header,section)})});
 const rels=await r.zip.get('ppt/slides/_rels/slide1.xml.rels').text();assert.match(rels,/Target="https:\/\/example.invalid\/first"/);assert.match(rels,/relationships\/slide" Target="slide1.xml"/);assert.match(rels,/mailto:person@example.invalid/);assert.equal(r.stats.drawings.hyperlinks,3);assert.ok(!r.losses.some(l=>l.id==='legacy-presentation-link'));
});

const footer=(id,type,field)=>shape(id,202,{text:'*',textRecords:[rec(field,u32(0))],client:[rec(3011,concat(u32(id-1024),Buffer.from([type,0,0,0])))]});
const hf=(flags,text,instance=3)=>container(4057,[rec(4058,concat(u16(0),u16(flags))),...(text==null?[]:[rec(4026,Buffer.from(text,'utf16le'),2)])],instance);
const transition=(type,direction,flags=1,speed=2,time=0)=>rec(1017,concat(u32(time),u32(0),Buffer.from([direction,type]),u16(flags),Buffer.from([speed,0,0,0])));
test('PPT legacy footer and slide number use their master anchors on the slide',async()=>{
 const r=await convert(make([shape(2050,202,{text:'Body'})],{masterRecords:[rec(1036,container(0xf002,[footer(1027,9,4090),footer(1028,8,4056)]))],documentRecords:[hf(40,'Confidential')]}));
 assert.match(r.xml,/<p:ph type="ftr"/);assert.match(r.xml,/<a:t>Confidential<\/a:t>/);assert.match(r.xml,/type="slidenum"/);assert.equal(r.stats.drawings.headerFields,2);assert.ok(!r.losses.some(l=>l.id==='legacy-slide-headers-footers'));assert.match(await r.zip.get('ppt/slideMasters/slideMaster1.xml').text(),/<p:hf hdr="0" ftr="1" dt="0" sldNum="1"/);
});
test('PPT per-slide footer text overrides the document footer',async()=>{
 const r=await convert(make([],{masterRecords:[rec(1036,container(0xf002,[footer(1027,9,4090)]))],documentRecords:[hf(32,'Global')],slideRecords:[hf(32,'Local',0)]}));assert.match(r.xml,/<a:t>Local<\/a:t>/);assert.doesNotMatch(r.xml,/>Global</);
});
test('PPT empty fixed header text does not add a date or a spurious loss',async()=>{
 const r=await convert(make([],{masterRecords:[rec(1036,container(0xf002,[footer(1027,7,4088)]))],documentRecords:[hf(5)]}));assert.doesNotMatch(r.xml,/<a:fld/);assert.ok(!r.losses.some(l=>l.id==='legacy-slide-headers-footers'));
});
test('PPT checker transition keeps orientation, speed, manual and automatic advance',async()=>{
 const r=await convert(make([],{slideRecords:[transition(3,0,1024,1,4321)]}));assert.match(r.xml,/<p:transition spd="med" advClick="0" advTm="4321"><p:checker dir="horz"\/>/);assert.equal(r.stats.drawings.transitions,1);assert.ok(!r.losses.some(l=>l.id==='legacy-slide-transition'));
});
test('PPT cut through black and stop-sound keep their separate effects',async()=>{
 const r=await convert(make([],{slideRecords:[transition(0,1,257)]}));assert.match(r.xml,/<p:transition><p:cut thruBlk="1"\/><p:sndAc><p:endSnd\/><\/p:sndAc>/);
});
test('PPT a transition sound loss does not drop its wheel or advance settings',async()=>{
 const r=await convert(make([],{slideRecords:[transition(26,8,1041,0,7000)]}));assert.match(r.xml,/<p:wheel spokes="8"\/>/);assert.match(r.xml,/advTm="7000"/);assert.ok(r.losses.some(l=>l.id==='legacy-presentation-sound'&&l.notify));
});

test('PPT WordArt letters stored outside textbox atoms remain editable text',async()=>{
 const str=Buffer.from('ART\r\0','utf16le'),face=Buffer.from('Arial Black\0','utf16le'),properties=rec(0xf00b,concat(u16(0x8000|192),u32(str.length),u16(0x8000|197),u32(face.length),u16(195),u32(48*65536),u16(255),u32(0x40004020),str,face),4,3);
 const s=container(0xf004,[rec(0xf00a,concat(u32(2050),u32(0xa00)),136,2),rec(0xf010,concat(...[0,0,2000,1000].map(u32))),properties]);
 const r=await convert(make([s]));assert.match(r.xml,/<a:prstTxWarp prst="textPlain"/);assert.match(r.xml,/<a:t>ART<\/a:t>/);assert.match(r.xml,/<a:latin typeface="Arial Black"/);assert.match(r.xml,/sz="4800" b="1"/);assert.equal(r.stats.drawings.wordArt,1);assert.equal((r.xml.match(/<a:p>/g)||[]).length,1,'the terminal paragraph mark must not add an empty warped line');
});

test('PPT a retained layout keeps its image dependency and uses the output master',async()=>{
 const png=await B.images.png(1,1,Uint8Array.from([3,4,5,255])),layout=B.xmlDeclaration+'<p:sldLayout xmlns:p="'+B.ns.p+'" xmlns:a="'+B.ns.a+'" xmlns:r="'+B.ns.officeRel.slice(0,-1)+'" type="blank"><p:cSld><p:spTree><p:pic><p:nvPicPr><p:cNvPr id="10" name="Picture"><a:hlinkClick r:id="" action="ppaction://hlinkshowjump?jump=nextslide"/></p:cNvPr><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic></p:spTree></p:cSld></p:sldLayout>';
 const parts=[{name:'ppt/slideLayouts/slideLayout7.xml',data:layout},{name:'ppt/slideLayouts/_rels/slideLayout7.xml.rels',data:B.el('Relationships',{xmlns:B.ns.rel},B.el('Relationship',{Id:'rId1',Type:B.ns.officeRel+'slideMaster',Target:'../slideMasters/slideMaster99.xml'})+B.el('Relationship',{Id:'rId2',Type:B.ns.officeRel+'image',Target:'../media/image7.png'}))},{name:'ppt/media/image7.png',data:png},{name:'[Content_Types].xml',data:B.el('Types',{xmlns:B.ns.content},B.el('Default',{Extension:'png',ContentType:'image/png'}))}];
 const zip=new Uint8Array(await(await L.zip.write(parts)).arrayBuffer());const r=await convert(make([],{masterRecords:[rec(1052,u32(1)),rec(1054,zip,1)],slideRecords:[rec(1058,concat(u32(1),u32(1)))]}));
 assert.equal(r.stats.storedLayouts,1);assert.match(await r.zip.get('ppt/slideLayouts/slideLayout1.xml').text(),/r:embed="rId2"/);const rels=await r.zip.get('ppt/slideLayouts/_rels/slideLayout1.xml.rels').text();assert.match(rels,/Id="rId1"[^>]+Target="..\/slideMasters\/slideMaster1.xml"/);assert.match(rels,/Id="rId2"[^>]+Target="..\/media\/image1.png"/);assert.deepEqual(await r.zip.get('ppt/media/image1.png').bytes(),png);assert.ok(!r.losses.some(l=>l.id==='legacy-presentation-design'));
});
test('PPT a missing retained-layout image falls back locally without losing the slide text',async()=>{
 const xml=B.el('p:sldLayout',{'xmlns:p':B.ns.p,'xmlns:a':B.ns.a,'xmlns:r':B.ns.officeRel.slice(0,-1)},B.el('p:cSld',{},B.el('p:spTree',{},B.el('a:blip',{'r:embed':'rId1'}))));
 const zip=new Uint8Array(await(await L.zip.write([{name:'ppt/slideLayouts/slideLayout1.xml',data:xml}])).arrayBuffer());
 const r=await convert(make([shape(2050,202,{text:'Kept'})],{masterRecords:[rec(1054,zip,1)]}));assert.match(r.xml,/>Kept</);assert.equal(r.stats.storedLayouts,0);assert.ok(r.losses.some(l=>l.id==='legacy-presentation-design'&&l.notify));
});

test('PPT slide text keeps literal dollar signs and JavaScript replacement sequences',async()=>{
 const text="$$PICTURE$$ $& $` $' $1";
 const r=await convert(make([shape(2050,202,{text})]));
 let actual='';const visit=()=>{};visit.text=(value,node)=>{if(node?.name==='a:t')actual+=value;};B.checkXml(r.xml,visit);
 assert.equal(actual,text);
});

test('PPT group placeholders become ordinary layout placeholders, never placeholder groups',async()=>{
 const g=container(0xf003,[container(0xf004,[rec(0xf00a,concat(u32(1025),u32(1)),0,2),rec(0xf009,concat(...[0,0,2000,1000].map(u32))),rec(0xf010,concat(...[0,0,2000,1000].map(u32))),rec(0xf011,rec(3011,concat(u32(2),Buffer.from([19,1,0,0]))))]),shape(1026,1,{child:true,text:'First'}),shape(1027,1,{child:true,text:'Second'})]);
 const r=await convert(make([g])),layout=await r.zip.get('ppt/slideLayouts/slideLayout1.xml').text();
 assert.match(r.xml,/<p:grpSp>/);assert.match(r.xml,/>First</);assert.match(r.xml,/>Second</);
 for(const xml of [r.xml,layout])for(const group of xml.match(/<p:nvGrpSpPr>.*?<\/p:nvGrpSpPr>/gs)||[])assert.doesNotMatch(group,/<p:ph/);
 assert.doesNotMatch(layout,/<p:grpSp>/);assert.match(layout,/<p:ph type="obj" idx="2" sz="half"/);
 assert.match(layout,/<a:ext cx="3175000" cy="1587500"/);
});

test('PPT empty textboxes have a text body so Office still renders their slides',async()=>{
 const r=await convert(make([shape(2050,202),shape(2051,202,{text:'Visible'})],{masterRecords:[rec(1036,container(0xf002,[shape(1025,202)]))]}));
 for(const part of ['ppt/slides/slide1.xml','ppt/slideMasters/slideMaster1.xml']){
  const xml=await r.zip.get(part).text();assert.match(xml,new RegExp('<p:cNvPr id="'+(part.includes('/slides/')?2050:1025)+'"'));for(const shape of xml.match(/<p:sp>.*?<\/p:sp>/gs)||[])if(/txBox="1"/.test(shape))assert.match(shape,/<p:txBody>.*<a:p>/s);
 }
 assert.match(r.xml,/>Visible</);
});
