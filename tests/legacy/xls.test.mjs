import {test} from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import {load} from '../../tools/legacy/load.mjs';
const B=load(),X=B.xls;
if(process.env.XLS_WRITE_SOURCE)vm.runInThisContext(fs.readFileSync(process.env.XLS_WRITE_SOURCE,'utf8'));
if(process.env.XLS_BEFORE){for(const key of ['convert','styles','sheetMetadata','richText','theme','styleExtensions'])delete X[key];delete B.checkXml;delete B.zipPart;vm.runInThisContext(fs.readFileSync(process.env.XLS_BEFORE,'utf8'));}
import {words,dword,rec,str,bof,workbook,cfb} from './xls-fixtures.mjs';
async function convert(globals,records){return X.convert(cfb(workbook(globals,[{name:'Data',records}])));}
async function xmls(result){const z=await L.zip.read(result.pkg);return Object.fromEntries(await Promise.all([...z].filter(([n])=>n.endsWith('.xml')).map(async([n,e])=>[n,await e.text()])));}
const cell=(row,col,xf=0)=>words(row,col,xf);
const number=(row,col,n)=>{const value=Buffer.alloc(8);value.writeDoubleLE(n);return rec(0x203,Buffer.concat([cell(row,col),value]));};
const font=(name,weight=400)=>{const b=Buffer.alloc(14);b.writeUInt16LE(220);b.writeUInt16LE(0x7fff,4);b.writeUInt16LE(weight,6);return rec(0x31,Buffer.concat([b,Buffer.from([name.length,1]),Buffer.from(name,'utf16le')]));};
const xf=(font=0,style=false)=>{const b=Buffer.alloc(20);b.writeUInt16LE(font);b.writeUInt16LE(style?5:1,4);b.writeUInt16LE(0x20,6);b.writeUInt16LE(0xfc00,8);return rec(0xe0,b);};

test('XLS conversion keeps typed value records and formula caches',async()=>{
  const formula=Buffer.alloc(25);cell(2,0).copy(formula);formula.writeDoubleLE(42,6);formula.writeUInt16LE(3,20);formula.set([0x1e,42,0],22);
  const result=await convert([rec(0x22,words(1)),rec(0x60),rec(0xfc,Buffer.concat([dword(1),dword(1),str('A\x07_x0041_')]))],[
    number(0,0,1.25),rec(0x27e,Buffer.concat([cell(0,1),dword((1234<<2)|3)])),rec(0xbd,Buffer.concat([words(0,2,0),dword((3<<2)|2),words(0),dword((4<<2)|2),words(3)])),
    rec(0xfd,Buffer.concat([cell(1,0),dword(0)])),rec(0x205,Buffer.concat([cell(1,1),Buffer.from([1,0])])),rec(0x205,Buffer.concat([cell(1,2),Buffer.from([7,1])])),rec(0x201,cell(1,3)),rec(0xbe,words(1,4,0,0,5)),rec(6,formula),rec(0x204,Buffer.concat([cell(3,0),str('中文')]))
  ]),xml=await xmls(result),sheet=xml['xl/worksheets/sheet1.xml'];
  assert.equal(result.type,'xltx');assert.equal(result.stats.cells,12);assert.match(sheet,/<c r="A1"><v>1.25<\/v><\/c>/);assert.match(sheet,/<c r="B1"><v>12.34<\/v>/);assert.match(sheet,/<c r="D1"><v>4<\/v>/);
  assert.match(sheet,/<c r="B2" t="b"><v>1<\/v>/);assert.match(sheet,/<c r="C2" t="e"><v>#DIV\/0!<\/v>/);assert.match(sheet,/<c r="D2"\/>/);assert.match(sheet,/<c r="A3"><f>42<\/f><v>42<\/v>/);
  assert.match(xml['xl/sharedStrings.xml'],/A_x0007__x005F_x0041_/);assert.match(xml['xl/sharedStrings.xml'],/中文/);assert.match(xml['xl/workbook.xml'],/date1904="1"/);assert.ok(!result.losses.some(l=>l.id==='legacy-formulas'));assert.equal(result.stats.formulas.converted,1);
});

test('XLS styles preserve skipped font index 4, named styles, custom formats and palette colors',()=>{
  const custom=Buffer.alloc(20);custom.writeUInt16LE(5);custom.writeUInt16LE(164,2);custom.writeUInt16LE(1,4);custom.writeUInt16LE(0xfc00,8);custom.writeUInt32LE(1<<26,14);custom.writeUInt16LE(8|(9<<7),18);
  const book=X.read(workbook([font('First'),font('Second'),font('Third'),font('Fourth'),font('Fifth',700),xf(0,true),xf(4),rec(0xe0,custom),rec(0x41e,Buffer.concat([words(164),str('0.0000')])),rec(0x293,Buffer.concat([words(0),str('Normal source')])),rec(0x92,Buffer.concat([words(2),Buffer.from([17,34,51,0,255,255,255,0])]))],[])),styles=X.styles(book,[]);
  assert.equal(styles.fontIndex(5),4);assert.equal(styles.fonts[styles.fontIndex(4)].weight,700);assert.match(styles.xml,/formatCode="0.0000"/);assert.match(styles.xml,/name="Normal source"/);assert.match(styles.xml,/rgb="FF112233"/);assert.match(styles.xml,/patternType="solid"><fgColor indexed="8"/);
});

test('SST rich runs cross CONTINUE records and retain font runs and phonetic text',()=>{
  const first=Buffer.concat([dword(1),dword(1),words(4),Buffer.from([8]),words(2),Buffer.from('AB')]);
  const continuation=Buffer.concat([Buffer.from([1]),Buffer.from('中文','utf16le'),words(0,0,2,5)]);
  const item=X.sst(X.records(Buffer.concat([rec(0xfc,first),rec(0x3c,continuation)])),0,false,true)[0];assert.equal(item.value,'AB中文');assert.deepEqual(item.runs,[{at:0,font:0},{at:2,font:5}]);
  const phonetic=Buffer.alloc(24);phonetic.writeUInt16LE(1);phonetic.writeUInt16LE(20,2);phonetic.writeUInt16LE(1,8);phonetic.writeUInt16LE(2,10);phonetic.writeUInt16LE(2,12);phonetic.write('かな',14,'utf16le');phonetic.writeUInt16LE(2,22);
  const styles=X.styles({version:'BIFF8',globals:[],codepage:1200},[]),losses=[];const xml=X.richText({value:'漢字',runs:[],extension:phonetic},styles,losses,'Data!A1');assert.match(xml,/<rPh sb="0" eb="2"><t>かな<\/t><\/rPh>/);assert.equal(losses.length,0);
});

test('rows, columns, merges and frozen pane selections keep their BIFF meanings',async()=>{
  const row=Buffer.alloc(16);row.writeUInt16LE(1);row.writeUInt16LE(420,6);row.writeUInt16LE(0x63,12);
  const window=Buffer.alloc(14);window.writeUInt16LE(2|4|8|16|32|128|256);window.writeUInt16LE(90,12);
  const result=await convert([], [rec(0x208,row),rec(0x7d,words(0,2,3000,0,0x301)),rec(0x81,words(128)),rec(0x225,words(0,300)),rec(0xe5,words(1,2,2,0,1)),rec(0x23e,window),rec(0xa0,words(3,4)),rec(0x41,words(1,2,2,1,0)),rec(0x1d,Buffer.concat([Buffer.from([0]),words(2,1,0,1,2,2),Buffer.from([1,1])])),number(2,0,8)]),xml=(await xmls(result))['xl/worksheets/sheet1.xml'];
  assert.match(xml,/summaryBelow="0" summaryRight="1"/);assert.match(xml,/defaultRowHeight="15"/);assert.match(xml,/width="11.71875"[^>]*hidden="1" outlineLevel="3"/);assert.match(xml,/<row r="2" ht="21" customHeight="1" hidden="1" outlineLevel="3"/);
  assert.match(xml,/zoomScale="75"/);assert.match(xml,/<pane xSplit="1" ySplit="2" topLeftCell="B3" activePane="bottomRight" state="frozen"/);assert.match(xml,/pane="bottomRight" activeCell="B3"/);assert.match(xml,/mergeCell ref="A3:B3"/);
});

test('BIFF5 RSTRING retains OEM text and two-byte formatting runs',()=>{
  const body=workbook([rec(0x42,words(437))],[{name:'A',records:[rec(0xd6,Buffer.concat([cell(0,0),words(2),Buffer.from([0x82,0x41,1,0,4])]))]}]);
  // Convert just the BoundSheet string to the BIFF5 single-byte representation.
  const records=X.records(body);for(const r of records)if(r.id===0x809)r.r.view.setUint16(0,0x500,true);
  const boundRec=records.find(r=>r.id===0x85);boundRec.r.bytes[7]=65;
  const book=X.read(body,{rich:true});assert.equal(book.sheets[0].cells[0].value,'éA');assert.deepEqual(book.sheets[0].cells[0].rich.runs,[{at:0,font:4}]);
});

test('format extensions keep theme index order, RGB color, gradient and font scheme',()=>{
  const color=Buffer.alloc(16);color.writeUInt16LE(3);color.writeUInt32LE(0,4);
  const ext=(type,data)=>Buffer.concat([words(type,data.length+4),data]);const header=Buffer.alloc(20);header.writeUInt16LE(0x87d);header.writeUInt16LE(1,14);header.writeUInt16LE(3,18);
  const gradient=Buffer.alloc(48+44);gradient.writeDoubleLE(45,4);gradient.writeUInt32LE(2,44);for(let i=0;i<2;i++){const p=48+i*22;gradient.writeUInt16LE(2,p);gradient.writeUInt32LE(i?0xffffffff:0xff332211,p+2);gradient.writeDoubleLE(i,p+6);}
  const book=X.read(workbook([font('Calibri'),xf(0,true),xf(0),rec(0x87d,Buffer.concat([header,ext(13,color),ext(14,Buffer.from([2])),ext(6,gradient)]))],[]));book.theme=true;const styles=X.styles(book,[]);
  assert.match(styles.xml,/<color theme="0"/);assert.match(styles.xml,/<scheme val="minor"/);assert.match(styles.xml,/<gradientFill type="linear" degree="45">/);assert.match(styles.xml,/rgb="FF112233"/);
});

test('bad sheet offsets, unfinished streams and overlapping merges fail cleanly',async()=>{
  await assert.rejects(()=>X.convert(rec(0x209,words(0,16))),{name:'LegacyError',code:'XLS_VERSION'});
  const bytes=workbook([],[{name:'A',records:[number(0,0,1)]}]);assert.throws(()=>X.read(bytes.subarray(0,bytes.length-4)),{name:'LegacyError',code:'SHEET_TRUNCATED'});
  const at=X.records(bytes).find(r=>r.id===0x85);at.r.view.setUint32(0,at.offset,true);assert.throws(()=>X.read(bytes),{name:'LegacyError',code:'SHEET_OFFSET'});
  await assert.rejects(()=>convert([],[rec(0xe5,words(2,0,2,0,2,1,3,1,3))]),{name:'LegacyError',code:'MERGE_OVERLAP'});
});

test('embedded theme XML and ZIP expansion have independent bounded checks',async()=>{
  assert.equal(B.checkXml('<a:x xmlns:a="urn:x" n="a&amp;b"><a:y/></a:x>').attrs.n,'a&b');
  for(const xml of ['<!DOCTYPE x><x/>','<x a="&bad;"/>','<x a="&constructor;"/>','<x><y/></z>','<a:x/>','<x/><x/>','<x a="&#0;"/>','<x xmlns:a="urn:x" xmlns:b="urn:x" a:n="1" b:n="2"/>'])assert.throws(()=>B.checkXml(xml),{name:'LegacyError'});
  const data=Buffer.alloc(100000,65),blob=await L.zip.write([{name:'theme.xml',data}]),bytes=new Uint8Array(await blob.arrayBuffer()),entry=(await L.zip.read(bytes)).get('theme.xml');
  await assert.rejects(()=>B.zipPart(bytes,entry,1000),{name:'LegacyError',code:'ZIP_PART_LIMIT'});await assert.rejects(()=>B.zipPart(bytes,{...entry,usize:2},200000),{name:'LegacyError'});assert.equal((await B.zipPart(bytes,entry)).length,data.length);
});

test('long cell text is shortened at the Office limit with an explicit content-loss notice',()=>{
  const styles=X.styles({version:'BIFF8',globals:[],codepage:1200},[]),losses=[];
  const xml=X.richText({value:'A'.repeat(32766)+'😀tail'},styles,losses,'Data!A1');
  assert.equal(xml,'<t>'+'A'.repeat(32766)+'</t>');assert.equal(losses.length,1);assert.equal(losses[0].notify,true);assert.equal(losses[0].place,'Data!A1');
  const exact='B'.repeat(32767);assert.equal(X.richText({value:exact},styles,[], 'Data!B1'),'<t>'+exact+'</t>');
});
