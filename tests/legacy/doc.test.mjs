import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import {load} from '../../tools/legacy/load.mjs';
import {doc,sprm,style,variable,u16,u32} from './doc-fixtures.mjs';
const B=load(),D=B.doc;
if(process.env.LEGACY_BEFORE_E1_AUDIT)for(const file of ['doc-sprm','doc-style','doc-structure','doc-lists','doc-xml','doc-settings','doc-text','doc-write'])vm.runInThisContext(fs.readFileSync(process.env.LEGACY_BEFORE_E1_AUDIT+'/'+file+'.js','utf8'));
if(process.env.LEGACY_BEFORE_E1){delete D.convert;const file=process.env.LEGACY_BEFORE_E1;vm.runInThisContext(fs.readFileSync(file,'utf8'),{filename:file});}
async function convert(bytes,options){const result=await D.convert(bytes,options),zip=await globalThis.L.zip.read(result.pkg);return {...result,xml:async name=>zip.get('word/'+name+'.xml').text()};}

test('Word text is packaged with its font, template type and Word 2003 compatibility',async()=>{
 const r=await convert(doc({text:'Café 漢字 😀\r',template:true}));
 assert.equal(r.type,'dotx');assert.match(await r.xml('document'),/Café 漢字 😀/);assert.match(await r.xml('fontTable'),/Times New Roman/);assert.match(await r.xml('settings'),/w:name="compatibilityMode"[^>]*w:val="11"/);
});
test('character positions follow out-of-order compressed and Unicode pieces, not physical offsets',async()=>{
 const pieces=[{text:'First ',compressed:true,offset:2500},{text:'漢字\r',offset:2048}];
 const r=await convert(doc({text:'First 漢字\r',pieces,chpx:[{start:2048,end:2054,ops:[sprm(0x0836,1)]},{start:2500,end:2506,ops:[sprm(0x0835,1)]}]}));
 const s=await r.xml('document');assert.match(s,/<w:b w:val="1"[^]*First /);assert.match(s,/<w:i w:val="1"[^]*漢字/);assert.ok(s.indexOf('First ')<s.indexOf('漢字'));
});
test('paragraph-mark properties, inherited styles and piece PRM toggles retain the source formatting',async()=>{
 const r=await convert(doc({text:'Heading\r',pieces:[{text:'Heading\r',offset:2048,prm:1}],prcs:[sprm(0x0835,129)],papx:[{start:2048,end:2064,style:1,ops:[sprm(0x2403,1)]}]}));
 const s=await r.xml('document');assert.match(s,/w:pStyle w:val="Heading1"/);assert.match(s,/w:jc w:val="center"/);assert.match(s,/w:b w:val="0"/);assert.match(await r.xml('styles'),/w:basedOn w:val="Normal"/);
});
test('section breaks, first/even/default headers and the header guard preserve story boundaries',async()=>{
 const headers=['','','','','','','Even\r\r','Odd\r\r','Even footer\r\r','Odd footer\r\r','First\r\r','First footer\r\r'],cp=[0];for(const h of headers)cp.push(cp.at(-1)+h.length);
 const dop=Buffer.alloc(616);dop.writeUInt16LE(1,0);
 const r=await convert(doc({text:'One\x0cTwo\r',stories:{headers:headers.join('')},headers:cp,dop,sections:[{end:4,ops:[sprm(0x300a,1)]},{end:8,ops:[sprm(0x3009,0),sprm(0xb01f,16838),sprm(0xb020,11906),sprm(0x301d,1)]}]}));
 const s=await r.xml('document');assert.equal((s.match(/<w:sectPr>/g)||[]).length,2);assert.match(s,/w:type w:val="continuous"/);assert.match(s,/w:headerReference w:type="first"/);assert.match(s,/w:pgSz w:w="16838" w:h="11906"/);assert.equal((await r.xml('header1')).match(/<w:p>/g).length,1);assert.match(await r.xml('settings'),/evenAndOddHeaders/);
});
test('footnote and endnote references resolve to their own stories',async()=>{
 const r=await convert(doc({text:'A\x02 B\x02\r',stories:{footnotes:'\x02 Foot note\r',endnotes:'\x02 End note\r'},footnotes:[{cp:1,start:0}],endnotes:[{cp:4,start:0}]}));
 assert.match(await r.xml('document'),/w:footnoteReference w:id="1"/);assert.match(await r.xml('document'),/w:endnoteReference w:id="1"/);assert.match(await r.xml('footnotes'),/w:footnoteRef/);assert.match(await r.xml('footnotes'),/Foot note/);assert.match(await r.xml('endnotes'),/End note/);
});
test('nested field instructions stay hidden while cached results and table-cell text remain',async()=>{
 const r=await convert(doc({text:'Start \x13SECRET \x13HIDDEN\x14ignored\x15\x14visible\x15\rCell one\x07Cell two\x07'}));
 const s=await r.xml('document'),fields=[];let visible='';
 const visit=node=>{if(node.name==='w:fldChar'){const type=node.attrs['w:fldCharType'];if(type==='begin')fields.push(false);else if(type==='separate')fields[fields.length-1]=true;else if(type==='end')fields.pop();}};
 visit.text=(text,node)=>{if(node?.name==='w:t'&&fields.every(Boolean))visible+=text;};B.checkXml(s,visit);
 assert.equal(visible,'Start visibleCell oneCell two');assert.match(s,/SECRET/);assert.match(s,/HIDDEN/);assert.ok(!r.losses.some(l=>l.id==='legacy-word-fields'));assert.ok(r.losses.some(l=>l.id==='legacy-word-tables'&&l.notify));
});
test('damaged optional formatting does not discard readable text; style cycles are bounded',async()=>{
 const r=await convert(doc({text:'Readable\r',styles:[style('Normal',1,[],[],1,0),style('Loop',1,[],[],0)],mutate:(word,table,pairs)=>{const [offset]=pairs.get(12);table.writeUInt32LE(0x3fffff,offset+8);}}));
 assert.match(await r.xml('document'),/Readable/);assert.ok(r.losses.some(l=>l.id==='legacy-word-damaged'));assert.ok(r.pkg.length);
});
test('old and encrypted Word files produce typed version/password refusals',async()=>{
 await assert.rejects(D.convert(doc({version:0x65})),e=>e.code==='DOC_VERSION');await assert.rejects(D.convert(doc({encrypted:true})),e=>e.code==='PASSWORD_REQUIRED');await assert.rejects(D.convert(doc({encrypted:true}),{password:'x'}),e=>e.code==='ENCRYPTION_UNSUPPORTED');
});
test('list LVLs outside lcbPlfLst and LFO start overrides retain editable numbering',async()=>{
 const lst=Buffer.alloc(30);lst.writeUInt16LE(1,0);lst.writeUInt32LE(0x12345678,2);lst[28]=1;for(let i=0;i<9;i++)lst.writeUInt16LE(4095,10+i*2);
 const lvl=Buffer.alloc(28);lvl.writeInt32LE(1,0);lvl[6]=1;const lv=Buffer.concat([lvl,u16(2),Buffer.from('\0.','utf16le')]);
 const lfo=Buffer.alloc(32);lfo.writeUInt32LE(1,0);lfo.writeUInt32LE(0x12345678,4);lfo[16]=1;lfo.writeInt32LE(7,24);lfo.writeUInt32LE(16,28);
 const r=await convert(doc({text:'List item\r',tables:[[73,Buffer.concat([lst,lv])],[74,lfo]],papx:[{start:2048,end:2068,ops:[sprm(0x460b,1),sprm(0x260a,0)]}],mutate:word=>word.writeUInt32LE(30,158+73*8)}));
 assert.match(await r.xml('document'),/w:numId w:val="1"/);assert.match(await r.xml('numbering'),/w:lvlText w:val="%1\."/);assert.match(await r.xml('numbering'),/w:startOverride w:val="7"/);
});
test('DOP compatibility reads the newer copy and reverses negative binary options',async()=>{
 const dop=Buffer.alloc(616);dop.writeUInt32LE(1<<13,84);dop.writeUInt32LE(1,508);dop.writeUInt32LE(1<<3,512);
 const r=await convert(doc({dop})),s=await r.xml('settings');assert.match(s,/w:noTabHangInd/);assert.match(s,/w:doNotExpandShiftReturn/);assert.doesNotMatch(s,/w:adjustLineHeightInTable/);
});
test('UTF-16 pieces retain a text BOM and a surrogate pair split by piece or formatting boundaries',async()=>{
 const pieces=[{text:'\ufeffA\ud83d',offset:2048},{text:'\ude00Z\r',offset:2500}];
 const r=await convert(doc({text:'\ufeffA😀Z\r',pieces,chpx:[{start:2048,end:2054,ops:[sprm(0x0835,1)]},{start:2500,end:2506,ops:[]}]}));
 assert.match(await r.xml('document'),/\ufeffA😀/);assert.doesNotMatch(await r.xml('document'),/�/);
});
test('grid exceptions and extended section page/footnote settings survive binary property variants',async()=>{
 const r=await convert(doc({text:'Grid\x02\r',stories:{footnotes:'\x02 Note\r'},footnotes:[{cp:4,start:0}],chpx:[{start:2048,end:2060,ops:[sprm(0x0868,0)]}],sections:[{end:6,ops:[sprm(0x3011,1),sprm(0x501c,2),sprm(0x7044,40000),sprm(0x303b,1),sprm(0x303c,2)]}]}));
 const s=await r.xml('document');assert.match(s,/w:snapToGrid w:val="0"/);assert.match(s,/w:pgNumType w:start="40000"/);assert.match(s,/w:pos w:val="beneathText"/);assert.match(s,/w:numRestart w:val="eachPage"/);
});
test('binary style highlighting and script flags are carried by runs, outside OOXML style definitions',async()=>{
 const r=await convert(doc({styles:[style('Normal',1,[],[sprm(0x2a0c,7),sprm(0x085a,1),sprm(0x0882,1)],4095,0)]}));
 assert.doesNotMatch(await r.xml('styles'),/<w:(?:highlight|rtl|cs)\b/);assert.match(await r.xml('document'),/<w:highlight w:val="yellow"/);assert.match(await r.xml('document'),/<w:rtl w:val="1"/);assert.match(await r.xml('document'),/<w:cs w:val="1"/);
});
test('style and plain resets respect property order and preserve highlighting and deleted text',async()=>{
 const r=await convert(doc({text:'Kept\rGone\r',styles:[style('Normal',1,[],[],4095,0),style('Emphasis',2,[],[sprm(0x0836,1)])],chpx:[{start:2048,end:2058,ops:[sprm(0x0835,1),sprm(0x2a0c,7),sprm(0x4a30,1),sprm(0x4a43,30)]},{start:2058,end:2068,ops:[sprm(0x0800,1),sprm(0x2a33,0)]}]}));
 const s=await r.xml('document');assert.match(s,/Kept/);assert.match(s,/<w:delText[^>]*>Gone<\/w:delText>/);assert.doesNotMatch(s,/<w:t[^>]*>Gone|<w:b\b/);assert.match(s,/<w:rStyle w:val="Emphasis"/);assert.match(s,/<w:sz w:val="30"/);assert.match(s,/<w:highlight w:val="yellow"/);
});
test('extended shading indices and border sides follow the WordprocessingML values and sequence',async()=>{
 const shd=Buffer.concat([u32(0),u32(0xffffff),u16(44)]),border=0x0108;
 const r=await convert(doc({papx:[{start:2048,end:2072,ops:[sprm(0x6427,border),sprm(0x6424,border),variable(0xc64d,shd)]}],sections:[{end:12,ops:[sprm(0x702e,border),sprm(0x702b,border)]}]}));
 const s=await r.xml('document');assert.match(s,/<w:shd w:val="pct37"/);assert.match(s,/<w:pBdr><w:top[^>]*\/><w:right/);assert.match(s,/<w:pgBorders[^>]*><w:top[^>]*\/><w:right/);
});
test('manual line breaks retain the side on which text clears an object',async()=>{
 const r=await convert(doc({text:'First\x0bLast\r',chpx:[{start:2048,end:2070,ops:[sprm(0x2879,3)]}]}));
 assert.match(await r.xml('document'),/<w:br w:clear="all"/);
});
