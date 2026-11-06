import {test}from 'node:test';import assert from 'node:assert/strict';import fs from 'node:fs';import vm from 'node:vm';
import {load}from '../../tools/legacy/load.mjs';import {rec,container,u16,u32,scheme,slideAtom,presentation}from './ppt-fixtures.mjs';
const B=load(),P=B.ppt;
if(process.env.LEGACY_BEFORE_D1){vm.runInThisContext(fs.readFileSync(process.env.LEGACY_BEFORE_D1,'utf8'));for(const key of ['structure','convert','write','prepareDesigns','roundTrip'])delete P[key];}
const parts=async bytes=>{const z=await L.zip.read(bytes),o={};for(const e of z.values())o[e.name]=await e.text();return o;};
const colors=['010203','111213','212223','313233','414243','515253','616263','717273'];

test('PowerPoint ignores an unreadable obsolete persist object',()=>{
 const f=presentation({stale:true}),book=P.read(f.fake);assert.equal(book.slides.length,2);assert.equal(book.slides[0].id,257);
});
test('PowerPoint structure keeps source order, units, fonts and notes ID namespaces',()=>{
 const book=P.structure(presentation().fake,[]);assert.deepEqual(book.slides.map(s=>s.id),[257,256]);assert.deepEqual(book.size,{cx:9144000,cy:6858000});assert.deepEqual(book.notesSize,{cx:6858000,cy:9144000});assert.equal(book.firstSlideNumber,5);assert.equal(book.rtl,true);assert.equal(book.slides[0].notes.id,300);assert.equal(book.slides[0].notes.slide,book.slides[0]);assert.equal(book.fonts.get(0).name,'Arial');
});
test('PowerPoint active schemes and title-master inheritance are separate from gallery schemes',async()=>{
 const pages={masters:[{id:0x80000003,ref:2,record:container(1016,[slideAtom(),scheme(P.defaultScheme,6),scheme(colors)])},{id:0x80000004,ref:3,record:container(1006,[slideAtom(0x80000003,0,7,2),scheme(P.defaultScheme)])}],slides:[{id:256,ref:4,record:container(1006,[slideAtom(0x80000004,0,7,0),scheme(P.defaultScheme)])}],notes:[]};
 const book=P.structure(presentation({pages}).fake,[]);assert.deepEqual(book.slides[0].scheme,colors);await P.prepareDesigns(book,book.losses);assert.equal(book.outputMasters.length,1);assert.equal(book.slides[0].outputLayout.type,'title');assert.equal(book.slides[0].outputLayout.master.source.id,0x80000003);
 pages.masters[0].record=container(1006,[slideAtom(0x80000004)]);const cyclic=P.structure(presentation({pages}).fake,[]);await P.prepareDesigns(cyclic,cyclic.losses);assert.ok(cyclic.losses.some(l=>l.id==='legacy-master-cycle'));assert.ok(cyclic.outputMasters.length);
});
test('PowerPoint structure does not fail for a damaged optional drawing container',()=>{
 const broken=Buffer.alloc(8);broken.writeUInt16LE(0xf00d,2);broken.writeUInt32LE(0xfffffff0,4);
 const pages={masters:[],slides:[{id:256,ref:2,record:container(1006,[slideAtom(),rec(1036,broken),rec(4008,Buffer.from('Still readable'))])}],notes:[]};
 const book=P.structure(presentation({pages}).fake,[]);assert.equal(book.slides.length,1);assert.deepEqual(book.slides[0].text,['Still readable']);assert.ok(book.losses.some(l=>l.notify));
});
test('PowerPoint packages have disjoint master/layout IDs and reciprocal notes relationships',async()=>{
 const result=await P.convert(presentation().bytes),p=await parts(result.pkg);assert.equal(result.type,'pptx');assert.equal(result.stats.slides,2);assert.equal(result.stats.notes,1);
 assert.match(p['ppt/presentation.xml'],/sldId id="257"[^>]*\/\><p:sldId id="256"/);assert.match(p['ppt/presentation.xml'],/firstSlideNum="5"/);
 const ids=Object.entries(p).filter(([n])=>n==='ppt/presentation.xml'||/^ppt\/slideMasters\/[^/]+xml$/.test(n)).flatMap(([,x])=>[...x.matchAll(/<p:sld(?:Master|Layout)Id id="(\d+)"/g)].map(m=>m[1]));assert.equal(ids.length,new Set(ids).size);
 assert.match(p['ppt/notesSlides/_rels/notesSlide1.xml.rels'],/slides\/slide1.xml/);assert.match(p['ppt/slides/_rels/slide1.xml.rels'],/notesSlides\/notesSlide1.xml/);assert.match(p['ppt/notesMasters/notesMaster1.xml'],/p:notesStyle/);
 assert.ok(result.losses.some(l=>l.id==='legacy-presentation-text'&&l.notify));assert.ok(!result.losses.some(l=>l.id.startsWith('legacy-propert')&&l.notify));
 for(const [name,xml]of Object.entries(p))if(name.endsWith('.xml'))B.checkXml(xml);
});
test('PowerPoint shows and templates use an explicit extension hint',async()=>{
 const f=presentation();for(const [extension,type,kind]of [['.pps','ppsx','slideshow'],['POT','potx','template']]){const r=await P.convert(f.bytes,{extension}),p=await parts(r.pkg);assert.equal(r.type,type);assert.match(p['[Content_Types].xml'],new RegExp('presentationml\\.'+kind+'\\.main'));}
});
test('PowerPoint keeps embedded themes and restores merged layout/master grouping',async()=>{
 const theme=B.xmlDeclaration+B.el('a:theme',{'xmlns:a':B.ns.a,name:'Original'},B.el('a:themeElements'));
 const layout=B.xmlDeclaration+B.el('p:sldLayout',{'xmlns:p':B.ns.p,type:'title'},B.el('p:cSld',{name:'Original title'}));
 const zip=async(name,xml)=>Buffer.from(await (await L.zip.write([{name,data:xml}])).arrayBuffer());
 const themeBytes=await zip('theme/theme1.xml',theme),layoutBytes=await zip('drs/slideLayouts/slideLayout1.xml',layout);
 const pages={masters:[{id:0x80000000,ref:2,record:container(1016,[slideAtom(),rec(1052,u32(1)),rec(1038,themeBytes)])},{id:0x80000001,ref:3,record:container(1016,[slideAtom(),rec(1053,u32(1)),rec(1054,layoutBytes,1)])}],slides:[{id:256,ref:4,record:container(1006,[slideAtom(0x80000001)])}],notes:[]};
 const book=P.structure(presentation({pages}).fake,[]);await P.prepareDesigns(book,book.losses);assert.equal(book.outputMasters.length,1);assert.equal(book.outputLayouts.length,1);assert.equal(book.outputLayouts[0].stored.xml,layout);assert.equal(book.outputMasters[0].source.theme.xml,theme);
 const f=presentation({pages}),r=await P.convert(f.bytes),p=await parts(r.pkg);assert.equal(p['ppt/theme/theme1.xml'],theme);assert.equal(p['ppt/slideLayouts/slideLayout1.xml'],layout);
});
test('PowerPoint damaged embedded designs are local losses, with independent native structure',async()=>{
 const pages={masters:[{id:0x80000000,ref:2,record:container(1016,[slideAtom(),scheme(colors),rec(1038,Buffer.from('bad zip')),rec(1054,Buffer.from('bad zip'),1)])}],slides:[{id:256,ref:3,record:container(1006,[slideAtom(0x80000000)])}],notes:[]};
 const result=await P.convert(presentation({pages}).bytes),p=await parts(result.pkg);assert.match(p['ppt/theme/theme1.xml'],/val="313233"/);assert.equal(result.stats.slides,1);assert.equal(result.losses.filter(l=>l.id==='legacy-presentation-design').length,2);assert.equal(result.losses.filter(l=>l.id==='legacy-presentation-design'&&l.notify).length,1);
});
test('PowerPoint tolerates a stale encrypted token on plain content and refuses invalid dimensions',async()=>{
 assert.equal((await P.convert(presentation({encrypted:true}).bytes)).stats.slides,2);await assert.rejects(P.convert(presentation({size:[0,4320,4320,5760]}).bytes),{code:'PPT_SIZE'});
});

test('PowerPoint retains local color-map overrides nested inside clrMapOvr',async()=>{
 const mapping=B.xmlDeclaration+B.el('a:clrMapOvr',{'xmlns:a':B.ns.a},B.el('a:overrideClrMapping',{...P.colorMap,bg1:'accent1',tx1:'lt1'}));
 const pages={masters:[],slides:[{id:256,ref:2,record:container(1006,[slideAtom(),rec(1039,Buffer.from(mapping))])}],notes:[]};
 const book=P.structure(presentation({pages}).fake,[]);await P.prepareDesigns(book,book.losses);assert.equal(book.slides[0].colorMap.bg1,'accent1');assert.equal(book.slides[0].colorMap.tx1,'lt1');
});

test('PowerPoint reports a lost picture background even without ordinary shapes',async()=>{
 const properties=rec(0xf00b,Buffer.concat([u16(384),u32(3),u16(390),u32(1)]),2,3),background=container(0xf004,[rec(0xf00a,Buffer.concat([u32(1),u32(0x400)]),1,2),properties]);
 const pages={masters:[],slides:[{id:256,ref:2,record:container(1006,[slideAtom(0,0,0,16),rec(1036,container(0xf002,[background]))])}],notes:[]};
 const r=await P.convert(presentation({pages}).bytes);assert.ok(r.losses.some(l=>l.id==='legacy-slide-background'&&l.notify));assert.ok(!r.losses.some(l=>l.id==='legacy-slide-drawings'));
});
