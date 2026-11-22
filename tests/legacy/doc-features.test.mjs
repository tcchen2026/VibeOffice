import {test} from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import {load} from '../../tools/legacy/load.mjs';
import {doc,u16,u32,sprm,variable} from './doc-fixtures.mjs';
const B=load(),D=B.doc;
if(process.env.LEGACY_BEFORE_E2)for(const name of ['doc-read','doc-sprm','doc-structure','doc-style','doc-lists','doc-xml','doc-settings','doc-text','doc-write'])vm.runInThisContext(fs.readFileSync(process.env.LEGACY_BEFORE_E2+'/'+name+'.js','utf8'));
const concat=(...parts)=>Buffer.concat(parts);
async function convert(bytes){const result=await D.convert(bytes),zip=await L.zip.read(result.pkg);return {...result,zip,xml:await zip.get('word/document.xml').text()};}
function fieldTable(text,flags=0){const events=Array.from({length:text.length},(_,cp)=>({ch:text.charCodeAt(cp),cp})).filter(x=>x.ch>=19&&x.ch<=21);return concat(...events.map(x=>u32(x.cp)),u32(text.length+1),...events.map(x=>Buffer.from([x.ch,x.ch===21?flags:0])));}
function texts(xml,tag){let out='';const visit=()=>{};visit.text=(t,node)=>{if(node?.name==='w:'+tag)out+=t;};B.checkXml(xml,visit);return out;}
const sttb=(names,extras=names.map(()=>Buffer.alloc(0)))=>concat(u16(65535),u16(names.length),u16(extras[0]?.length||0),...names.map((s,i)=>concat(u16(s.length),Buffer.from(s,'utf16le'),extras[i])));
const plc=(cp,records=[])=>concat(...cp.map(u32),...records);
const bookmarkTables=(text,names,starts,ends,indices)=>[[21,sttb(names)],[22,plc([...starts,text.length+1],indices.map(i=>concat(u16(i),u16(0x4000))))],[23,plc([...ends,text.length+1])]];
function tableDef(edges,flags=[]){const data=concat(Buffer.from([edges.length-1]),...edges.map(n=>u16(n&65535)),...flags.map(f=>concat(u16(f),u16(0),Buffer.alloc(16))));return concat(u16(0xd608),u16(data.length+1),data);}
function tableDoc(parts){let cp=0;const papx=parts.map(p=>{const start=2048+cp*2;cp+=p.text.length;return {start,end:2048+cp*2,ops:[...(p.depth?[sprm(0x2416,1),sprm(0x6649,p.depth)]:[]),...(p.row?[sprm(p.depth===1?0x2417:0x244c,1)]:[]),...(p.cell&&p.depth>1?[sprm(0x244b,1)]:[]),...(p.ops||[])]};});return doc({text:parts.map(p=>p.text).join(''),papx});}
const xst=s=>concat(u16(s.length),Buffer.from(s,'utf16le')),xstz=s=>concat(xst(s),u16(0));
function formDoc(kind,bits,extra,cache='cached'){
 const text='Before \x13'+['FORMTEXT','FORMCHECKBOX','FORMDROPDOWN'][kind]+' \x01\x14'+cache+'\x15 after\r',cp=text.indexOf('\x01');
 const ff=concat(u32(0xffffffff),u16(bits|kind),u16(kind?0:42),u16(28),xstz('Control'),extra),data=concat(u32(68+ff.length),u16(68),Buffer.alloc(62),ff);
 return {text,data,chpx:[{start:2048,end:2048+cp*2,ops:[]},{start:2048+cp*2,end:2048+(cp+1)*2,ops:[sprm(0x6a03,0),sprm(0x0806,1)]},{start:2048+(cp+1)*2,end:2048+text.length*2,ops:[]}]};
}

test('Word complex fields retain nested instructions, cached text and lock state',async()=>{
 const text='Before \x13IF \x13PAGE\x141\x15 = 1 "yes" "no"\x14yes\x15 after\r';
 const r=await convert(doc({text,tables:[[16,fieldTable(text,0x94)]]}));
 assert.equal((r.xml.match(/w:fldCharType="begin"/g)||[]).length,2);
 assert.equal((r.xml.match(/w:fldCharType="separate"/g)||[]).length,2);
 assert.equal((r.xml.match(/w:fldCharType="end"/g)||[]).length,2);
 assert.match(r.xml,/w:fldLock="1" w:dirty="1"/);assert.equal(texts(r.xml,'instrText'),'IF PAGE = 1 "yes" "no"');
 assert.equal(texts(r.xml,'t'),'Before 1yes after');assert.ok(!r.losses.some(x=>x.id==='legacy-word-fields'));
});
test('Word instruction-only index fields survive without a PLC field entry',async()=>{
 const r=await convert(doc({text:'Word\x13 XE "Index entry" \x15 text\r'}));
 assert.equal(texts(r.xml,'instrText'),' XE "Index entry" ');assert.equal(texts(r.xml,'t'),'Word text');
 assert.equal((r.xml.match(/w:fldCharType=/g)||[]).length,2);assert.ok(!r.losses.some(x=>x.id==='legacy-word-fields'));
});
test('Word field positions are local to their story and broken flag tables keep live fields',async()=>{
 const main='Body \x13PAGE\x147\x15\r',header='Header \x13DATE\x142020\x15\r\r';
 const r=await convert(doc({text:main,stories:{headers:header},headers:[0,0,0,0,0,0,0,header.length],tables:[[16,Buffer.from([1])],[17,fieldTable(header,16)]]}));
 assert.equal(texts(r.xml,'instrText'),'PAGE');assert.equal(texts(r.xml,'t'),'Body 7');
 const h=await r.zip.get('word/header1.xml').text();assert.equal(texts(h,'instrText'),'DATE');assert.match(h,/w:fldLock="1"/);
 assert.ok(r.losses.some(x=>x.id==='legacy-word-damaged'));assert.ok(!r.losses.some(x=>x.id==='legacy-word-fields'));
});
test('Word an incomplete field retains its cache without emitting unmatched delimiters',async()=>{
 const r=await convert(doc({text:'Prefix \x13BAD\x14cached\r'}));
 assert.equal(texts(r.xml,'t'),'Prefix cached');assert.doesNotMatch(r.xml,/w:fldChar/);
 assert.ok(r.losses.some(x=>x.id==='legacy-word-fields'&&x.notify));
});
test('Word text form fields retain defaults, format, limits, help and entry macros',async()=>{
 const options=formDoc(0,0x4080,concat(xstz('Default'),xstz('Uppercase'),xstz('Fill this in'),xstz(''),xstz('StartForm'),xstz(''))),r=await convert(doc(options));
 assert.match(r.xml,/<w:ffData>[^]*<w:calcOnExit w:val="1"/);assert.match(r.xml,/<w:entryMacro w:val="StartForm"/);assert.match(r.xml,/<w:helpText w:type="text" w:val="Fill this in"/);
 assert.match(r.xml,/<w:textInput><w:type w:val="regular"\/><w:default w:val="Default"\/><w:maxLength w:val="42"\/><w:format w:val="Uppercase"\/><\/w:textInput>/);
 assert.equal(texts(r.xml,'t'),'Before cached after');assert.equal(r.stats.formFields,1);assert.ok(!r.losses.some(x=>x.id==='legacy-word-picture'));
});
test('Word checkbox and drop-down form fields retain current and default choices',async()=>{
 const rest=concat(...Array.from({length:5},()=>xstz('')));
 const checked=await convert(doc(formDoc(1,0x0404,concat(u16(0),rest),'')));
 assert.match(checked.xml,/<w:checkBox><w:size w:val="28"\/><w:default w:val="0"\/><w:checked w:val="1"\/><\/w:checkBox>/);
 const dropdown=await convert(doc(formDoc(2,0x8004,concat(u16(0),rest,sttb(['First','Second'])),'Second')));
 assert.match(dropdown.xml,/<w:ddList><w:result w:val="1"\/><w:default w:val="0"\/><w:listEntry w:val="First"\/><w:listEntry w:val="Second"\/><\/w:ddList>/);
 assert.ok(!dropdown.losses.some(x=>x.notify));
});
test('Word damaged form data keeps the cached value and other live fields',async()=>{
 const options=formDoc(0,0,Buffer.from([255])),r=await convert(doc(options));
 assert.equal(texts(r.xml,'t'),'Before cached after');assert.match(r.xml,/w:fldCharType="begin"/);assert.doesNotMatch(r.xml,/<w:ffData>/);
 assert.ok(r.losses.some(x=>x.notify&&x.what.includes('form field')));assert.ok(!r.losses.some(x=>x.id==='legacy-word-picture'));
});
test('Word document variables retain values without exporting binary VBA signatures as text',async()=>{
 const names=['Customer','Year','Sign'],vars=concat(sttb(names,names.map(()=>u32(0))),xst('Acme & Co'),xst('2020'),xst('binary signature'));
 const r=await convert(doc({text:'\x13DOCVARIABLE Customer\x14Acme & Co\x15\r',tables:[[60,vars]]})),settings=await r.zip.get('word/settings.xml').text();
 assert.match(settings,/<w:docVars><w:docVar w:name="Customer" w:val="Acme &amp; Co"\/><w:docVar w:name="Year" w:val="2020"\/><\/w:docVars>/);assert.doesNotMatch(settings,/binary signature/);
});
test('Word fields cross paragraph boundaries without losing UTF-16 text or bookmark identities',async()=>{
 const text='\x13IF "😀" = "😀"\x14First\rSecond😀\x15\r',r=await convert(doc({text,tables:bookmarkTables(text,['Result'],[text.indexOf('First')],[text.indexOf('\x15')],[0])}));
 assert.equal(texts(r.xml,'t'),'FirstSecond😀');assert.equal(texts(r.xml,'instrText'),'IF "😀" = "😀"');assert.equal(r.stats.fields,1);assert.equal(r.stats.bookmarks,1);
});
test('Word bookmarks join shuffled end indices and split runs at UTF-16 boundaries',async()=>{
 const text='A😀BCD\r',r=await convert(doc({text,tables:bookmarkTables(text,['Outer','Inner','Point'],[0,3,5],[4,5,6],[2,0,1])}));
 const marks=[];let cp=0;const visit=node=>{if(node.name==='w:bookmarkStart'||node.name==='w:bookmarkEnd')marks.push([node.name,node.attrs['w:id'],cp,node.attrs['w:name']]);};visit.text=(s,node)=>{if(node?.name==='w:t')cp+=s.length;};B.checkXml(r.xml,visit);
 assert.deepEqual(marks,[['w:bookmarkStart','0',0,'Outer'],['w:bookmarkStart','1',3,'Inner'],['w:bookmarkEnd','1',4,undefined],['w:bookmarkStart','2',5,'Point'],['w:bookmarkEnd','2',5,undefined],['w:bookmarkEnd','0',6,undefined]]);
 assert.equal(r.stats.bookmarks,3);assert.ok(!r.losses.some(x=>x.id==='legacy-word-bookmarks'));
});
test('Word a broken bookmark end loses only that bookmark',async()=>{
 const text='Readable\r',r=await convert(doc({text,tables:bookmarkTables(text,['Good','Bad'],[0,1],[2,4],[0,9])}));
 assert.match(r.xml,/w:name="Good"/);assert.doesNotMatch(r.xml,/w:name="Bad"/);assert.equal(texts(r.xml,'t'),'Readable');assert.ok(r.losses.some(x=>x.id==='legacy-word-bookmarks'&&x.notify));
});
test('Word a damaged bookmark inside a surrogate pair does not corrupt the glyph',async()=>{
 const text='A😀B\r',r=await convert(doc({text,tables:bookmarkTables(text,['Good','Broken'],[0,2],[3,4],[0,1])}));
 assert.equal(texts(r.xml,'t'),'A😀B');assert.equal(r.stats.bookmarks,1);assert.doesNotMatch(r.xml,/w:name="Broken"/);assert.ok(r.losses.some(x=>x.id==='legacy-word-bookmarks'&&x.notify));
});
test('Word comments preserve text, author, range and point reference',async()=>{
 const text='AB\x05 CD\x05\r',body='\x05 One\r\x05 Two\r',a=Buffer.alloc(30),b=Buffer.alloc(30);
 a.writeUInt16LE(2,0);Buffer.from('AA','utf16le').copy(a,2);a.writeUInt32LE(42,26);b.writeUInt32LE(0xffffffff,26);
 const r=await convert(doc({text,stories:{comments:body},tables:[[4,plc([2,6,text.length+1],[a,b])],[5,plc([0,6,body.length])],[36,concat(u16(5),Buffer.from('Alice','utf16le'))],[37,sttb([''],[concat(u16(0x100),u32(42),u32(0xffffffff))])],[42,plc([0,text.length+1],[concat(u16(0),u16(0))])],[43,plc([2,text.length+1])]]}));
 assert.match(r.xml,/<w:commentRangeStart w:id="0"/);assert.match(r.xml,/<w:commentRangeEnd w:id="0"/);assert.equal((r.xml.match(/<w:commentReference /g)||[]).length,2);
 const c=await r.zip.get('word/comments.xml').text();assert.equal(texts(c,'t'),' One Two');assert.match(c,/w:author="Alice" w:initials="AA"/);assert.equal((c.match(/w:annotationRef/g)||[]).length,2);assert.equal(r.stats.comments,2);assert.ok(!r.losses.some(x=>x.notify));
});
test('Word insertion, deletion and paragraph-mark revisions retain author, date and deleted text',async()=>{
 const text='New\rOld\r',date=(120<<20)|(1<<16)|(2<<11)|(3<<6)|4,dop=Buffer.alloc(616);dop[5]=128;
 const r=await convert(doc({text,dop,tables:[[51,sttb(['Unknown','Alice','Bob'])]],chpx:[{start:2048,end:2056,ops:[sprm(0x0801,1),sprm(0x4804,1),sprm(0x6805,date)]},{start:2056,end:2064,ops:[sprm(0x0800,1),sprm(0x4863,2),sprm(0x6864,date),sprm(0x2a33,0)]}]}));
 assert.equal(texts(r.xml,'t'),'New');assert.equal(texts(r.xml,'delText'),'Old');assert.match(r.xml,/<w:ins w:id="\d+" w:author="Alice" w:date="2020-01-02T03:04:00Z"/);assert.match(r.xml,/<w:del w:id="\d+" w:author="Bob"/);assert.match(r.xml,/<w:pPr>[^]*?<w:rPr><w:del/);assert.match(await r.zip.get('word/settings.xml').text(),/w:trackRevisions/);assert.ok(!r.losses.some(x=>x.id==='legacy-word-revisions'));
});
test('Word formatting revisions retain the properties before the revision wall',async()=>{
 const prop=concat(Buffer.from([1]),u16(1),u32(0));
 const r=await convert(doc({text:'Changed\r',tables:[[51,sttb(['Unknown','Alice'])]],chpx:[{start:2048,end:2064,ops:[sprm(0x0835,0),sprm(0x2a83,1),sprm(0x0835,1),variable(0xca89,prop)]}],papx:[{start:2048,end:2064,ops:[sprm(0x2403,0),sprm(0x2664,1),sprm(0x2403,2),variable(0xc66f,prop)]}]}));
 assert.match(r.xml,/<w:b w:val="1"\/><w:rPrChange[^>]*><w:rPr><w:b w:val="0"/);assert.match(r.xml,/<w:jc w:val="right"/);assert.match(r.xml,/<w:pPrChange[^>]*><w:pPr><w:jc w:val="left"/);assert.equal(texts(r.xml,'t'),'Changed');assert.ok(!r.losses.some(x=>/revision/.test(x.id)));
});
test('Word table cells retain widths, header, borders, shading and padding',async()=>{
 const r=await convert(tableDoc([{text:'A\x07',depth:1},{text:'B\x07',depth:1},{text:'\x07',depth:1,row:true,ops:[tableDef([-108,1692,3492],[0,128]),sprm(0x3404,1),sprm(0x3466,1),sprm(0x9407,480),variable(0xd605,concat(...Array.from({length:6},()=>u32(0x010104)))),variable(0xd609,concat(u16(7<<5),u16(0))),variable(0xd634,Buffer.from([0,1,10,3,120,0]))]},{text:'After\r'}]));
 assert.equal(r.stats.tables,1);assert.equal(r.stats.rows,1);assert.equal(r.stats.cells,2);assert.match(r.xml,/<w:gridCol w:w="1800"\/><w:gridCol w:w="1800"/);assert.match(r.xml,/<w:tblHeader\/>/);assert.match(r.xml,/<w:cantSplit\/>/);assert.match(r.xml,/<w:trHeight w:val="480" w:hRule="atLeast"/);assert.match(r.xml,/<w:shd[^>]*w:fill="FFFF00"/);assert.match(r.xml,/<w:vAlign w:val="center"/);assert.equal(texts(r.xml,'t'),'ABAfter');assert.ok(!r.losses.some(x=>x.notify));
});
test('Word nested tables use inner row markers and retain the containing cell paragraph',async()=>{
 const r=await convert(tableDoc([{text:'Inner\r',depth:2,cell:true},{text:'\r',depth:2,row:true,ops:[tableDef([0,1200])]},{text:'Outer\x07',depth:1},{text:'\x07',depth:1,row:true,ops:[tableDef([0,3600])]},{text:'After\r'}]));
 assert.equal(r.stats.tables,2);assert.equal(r.stats.rows,2);assert.match(r.xml,/<w:tc><w:tcPr>[^]*?<w:tbl>[^]*?<\/w:tbl><w:p>/);assert.equal(texts(r.xml,'t'),'InnerOuterAfter');assert.ok(!r.losses.some(x=>x.id==='legacy-word-tables'));
});
test('Word table grid spans follow differing cell boundaries across rows',async()=>{
 const r=await convert(tableDoc([{text:'Wide\x07',depth:1},{text:'\x07',depth:1,row:true,ops:[tableDef([0,3600])]},{text:'Left\x07',depth:1},{text:'Right\x07',depth:1},{text:'\x07',depth:1,row:true,ops:[tableDef([0,1800,3600])]},{text:'\r'}]));
 assert.equal(r.stats.rows,2);assert.match(r.xml,/<w:gridSpan w:val="2"/);assert.equal(texts(r.xml,'t'),'WideLeftRight');
});
test('Word horizontal merges retain hidden source cell text for later unmerging',async()=>{
 const r=await convert(tableDoc([{text:'Visible\x07',depth:1},{text:'Hidden\x07',depth:1},{text:'\x07',depth:1,row:true,ops:[tableDef([0,1800,3600],[1,2])]},{text:'\r'}]));
 assert.match(r.xml,/<w:hMerge w:val="restart"/);assert.match(r.xml,/<w:hMerge w:val="continue"/);assert.equal(texts(r.xml,'t'),'VisibleHidden');assert.equal(r.stats.cells,2);
});
test('Word a damaged table retains every cell in source order',async()=>{
 const r=await convert(tableDoc([{text:'First\x07',depth:1},{text:'Second\x07',depth:1},{text:'\x07',depth:1,row:true,ops:[tableDef([0,0,3600])]},{text:'After\r'}]));
 assert.equal(texts(r.xml,'t'),'FirstSecondAfter');assert.doesNotMatch(r.xml,/<w:tbl>/);assert.ok(r.losses.some(x=>x.id==='legacy-word-tables'&&x.notify));
});
test('Word text in a damaged table row terminator survives exactly once',async()=>{
 const r=await convert(tableDoc([{text:'Cell\x07',depth:1},{text:'Unexpected\x07',depth:1,row:true,ops:[tableDef([0,3600])]},{text:'After\r'}]));
 assert.equal(texts(r.xml,'t'),'CellUnexpectedAfter');assert.doesNotMatch(r.xml,/<w:tbl>/);assert.ok(r.losses.some(x=>x.id==='legacy-word-tables'&&x.notify));
});
test('Word tracked table formatting retains previous cell and row properties',async()=>{
 const revision=variable(0xd667,concat(Buffer.from([1]),u16(0),u32(0))),r=await convert(tableDoc([{text:'Cell\x07',depth:1},{text:'\x07',depth:1,row:true,ops:[tableDef([0,3600],[0]),sprm(0x9407,200),sprm(0x3668,1),sprm(0x9407,400),variable(0xd62c,Buffer.from([0,1,2])),revision]},{text:'\r'}]));
 assert.match(r.xml,/<w:trHeight w:val="400"/);assert.match(r.xml,/<w:trPrChange[^>]*><w:trPr><w:trHeight w:val="200"/);
 assert.match(r.xml,/<w:vAlign w:val="bottom"\/><w:tcPrChange[^>]*><w:tcPr>/);assert.match(r.xml,/<w:tblPrChange/);assert.ok(!r.losses.some(x=>x.id.startsWith('legacy-word-revision')));
});
test('Word automatic COLORREF uses its flag byte, including black component bytes',async()=>{
 const shd=concat(u32(0xff000000),u32(0xff000000),u16(0)),r=await convert(tableDoc([{text:'Readable\x07',depth:1},{text:'\x07',depth:1,row:true,ops:[tableDef([0,3600]),variable(0xd670,shd)]},{text:'\r'}]));
 assert.match(r.xml,/<w:shd w:val="clear" w:color="auto" w:fill="auto"/);assert.equal(D.color(0x00332211),'112233');assert.equal(D.color(0xffccbbaa),'auto');
});
test('Word a displaced automatic note anchor preserves punctuation and attaches to its marker',async()=>{
 const r=await convert(doc({text:'A\x02).\r',stories:{footnotes:'\x02Note\r'},footnotes:[{cp:2,start:0}]}));
 assert.equal(texts(r.xml,'t'),'A).');assert.match(r.xml,/<w:t xml:space="preserve">A<\/w:t><w:footnoteReference w:id="1"\/><w:t xml:space="preserve">\)\.<\/w:t>/);
 assert.equal(texts(await r.zip.get('word/footnotes.xml').text(),'t'),'Note');assert.ok(!r.losses.some(x=>x.notify));
 const missing=await convert(doc({text:'A).\r',stories:{footnotes:'\x02Note\r'},footnotes:[{cp:1,start:0}]}));
 assert.equal(texts(missing.xml,'t'),'A).');assert.match(missing.xml,/<w:footnoteReference w:id="1"/);
 const atEnd=await convert(doc({text:'A).\r',stories:{footnotes:'\x02Note\r'},footnotes:[{cp:3,start:0}]}));
 assert.equal(texts(atEnd.xml,'t'),'A).');assert.match(atEnd.xml,/<w:footnoteReference w:id="1"/);
});
test('Word adjacent floating tables split when their anchor changes, including a following inline table',async()=>{
 const row=(text,x,y)=>[{text:text+'\x07',depth:1},{text:'\x07',depth:1,row:true,ops:[tableDef([0,1800]),...(y==null?[]:[sprm(0x360d,0x20),sprm(0x940e,x&65535),sprm(0x940f,y&65535)])]}];
 const r=await convert(tableDoc([...row('One',-4,-4),...row('Two',-4,-4),...row('Three',-232,1706),...row('Four'),{text:'After\r'}]));
 assert.equal(r.stats.tables,3);assert.equal(r.stats.rows,4);assert.equal(texts(r.xml,'t'),'OneTwoThreeFourAfter');
 assert.equal((r.xml.match(/<w:tblpPr /g)||[]).length,2);assert.match(r.xml,/w:tblpX="-233" w:tblpY="1705"/);assert.ok(!r.losses.some(x=>x.notify));
});
test('Word a repeated partial cell definition retains unspecified cells and separate padding',async()=>{
 const r=await convert(tableDoc([{text:'First\x07',depth:1},{text:'Second\x07',depth:1},{text:'\x07',depth:1,row:true,ops:[tableDef([0,1800,3600],[0,128]),variable(0xd620,concat(Buffer.from([0,2,4]),u32(0x010104))),variable(0xd632,Buffer.from([0,2,2,3,120,0])),tableDef([0,1800,3600],[0])]},{text:'\r'}]));
 const cells=[...r.xml.matchAll(/<w:tcPr>([^]*?)<\/w:tcPr>/g)].map(m=>m[1]);
 assert.match(cells[1],/<w:bottom w:val="single"/);assert.match(cells[1],/<w:vAlign w:val="center"/);
 assert.ok(cells.every(c=>c.includes('<w:left w:w="120" w:type="dxa"')));assert.equal(texts(r.xml,'t'),'FirstSecond');
});
