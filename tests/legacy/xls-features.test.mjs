import {test} from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import {load} from '../../tools/legacy/load.mjs';
import {words,dword,rec,str,workbook,cfb} from './xls-fixtures.mjs';
const B=load(),X=B.xls;
if(process.env.XLS_C3_PERCENT_BEFORE)vm.runInThisContext(fs.readFileSync(process.env.XLS_C3_PERCENT_BEFORE,'utf8'));
if(process.env.XLS_C3_THEME_BEFORE)vm.runInThisContext(fs.readFileSync(process.env.XLS_C3_THEME_BEFORE,'utf8'));
if(process.env.XLS_C3_BREAKS_BEFORE)vm.runInThisContext(fs.readFileSync(process.env.XLS_C3_BREAKS_BEFORE,'utf8'));
if(process.env.XLS_C3_SDK_BEFORE)for(const name of ['xls-sheet','xls-write','xls-style','xls-formula-book','xls-misc'])vm.runInThisContext(fs.readFileSync(process.env.XLS_C3_SDK_BEFORE+'/'+name+'.js','utf8'));
if(process.env.XLS_C3_AUDIT_BEFORE)for(const name of ['xls-rules','xls-comments','xls-dxf','package'])vm.runInThisContext(fs.readFileSync(process.env.XLS_C3_AUDIT_BEFORE+'/'+name+'.js','utf8'));
if(process.env.XLS_C3_BEFORE)for(const name of ['package','xls-sheet','xls-style-ext','xls-formula','xls-formula-book','xls-style','xls-write'])vm.runInThisContext(fs.readFileSync(process.env.XLS_C3_BEFORE+'/'+name+'.js','utf8'));
const cat=(...a)=>Buffer.concat(a.map(a=>Buffer.from(a))),u=n=>Buffer.from([n]),int=n=>cat([0x1e],words(n)),ref=(row,col)=>cat([0x24],words(row,col));
const range=(r1,c1,r2=r1,c2=c1)=>words(r1,r2,c1,c2),f64=n=>{const b=Buffer.alloc(8);b.writeDoubleLE(n);return b;};
async function convert(records,globals=[]){const out=await X.convert(cfb(workbook(globals,[{name:'Data',records}]))),zip=await L.zip.read(out.pkg);return {out,zip,xml:await zip.get('xl/worksheets/sheet1.xml').text(),styles:await zip.get('xl/styles.xml').text(),book:await zip.get('xl/workbook.xml').text()};}
const dv=(flags,one,two=Buffer.alloc(0),regions=[range(2,1,4,1)])=>rec(0x1be,cat(dword(flags),str('Pick'),str('Invalid'),str('Choose a value'),str('Try again'),words(one.length,0),one,words(two.length,0),two,words(regions.length),...regions));
const cfHeader=(id=1,regions=[range(2,1,4,1)])=>rec(0x1b0,cat(words(2,id<<1),regions[0],words(regions.length),...regions));
const cf=(one,{type=1,op=5,dxf=cat(dword(0x3fffff),words(0)),two=Buffer.alloc(0)}={})=>rec(0x1b1,cat([type,op],words(one.length,two.length),dxf,one,two));
const cfEx=(rule,priority,flags=1,id=1)=>rec(0x87b,cat(Buffer.alloc(12),dword(0),words(id,rule),[0,1],words(priority),[flags,0,16],Buffer.alloc(16)));
const cf12=(type,tail,priority=1)=>rec(0x87a,cat(Buffer.alloc(12),[type,0],words(0,0),dword(0),words(0,0),[1],words(priority,0),[16],Buffer.alloc(16),tail));
const color=rgb=>cat(dword(2),Buffer.from(rgb+'ff','hex'),f64(0));
const threshold=(type,value)=>cat([type],words(0),value==null?[]:f64(value));
const escher=(id,body,options=0)=>cat(words(options,id),dword(body.length),body);
function commentRecords({value='Hi文',badText=false}={}){
  const shapeId=2049,objectId=7,shapeHead=escher(0xf00a,cat(dword(shapeId),dword(0xa00)),202<<4|2),anchor=escher(0xf010,words(0,2,256,1,64,4,512,5,128));
  const prefix=cat(shapeHead,anchor,escher(0xf011,[])),textMarker=escher(0xf00d,[]),shape=escher(0xf004,cat(prefix,textMarker),15);
  const txo=Buffer.alloc(18);txo.writeUInt16LE(value.length,10);txo.writeUInt16LE(16,12);
  const fontRuns=cat(words(0,0),dword(0),words(value.length,0),dword(0));
  const records=[rec(0xec,shape.subarray(0,-8)),rec(0x5d,cat(words(0x15,18,25,objectId),Buffer.alloc(14))),rec(0xec,textMarker),rec(0x1b6,txo)];
  if(!badText)records.push(rec(0x3c,cat([0],Buffer.from(value.slice(0,2),'latin1'))),rec(0x3c,cat([1],Buffer.from(value.slice(2),'utf16le'))),rec(0x3c,fontRuns));
  records.push(rec(0x1c,cat(words(2,1,2,objectId),str('Author'))));return records;
}

test('sheet names quote only ambiguous Excel names, including external and 3-D qualifiers',()=>{
  for(const name of ['Sheet1','Données','Data_2','Data.2','XFE1','A1048577','Sheet1:Sheet3','[1]Sheet1'])assert.equal(X.quoteSheet(name),name);
  for(const name of ['A1','XFD1048576','R1C1','RC','R1data','TRUE','false','2Data','Two words','O\'Brien','Sheet1:Other sheet','[1]Other sheet'])assert.equal(X.quoteSheet(name),"'"+name.replace(/'/g,"''")+"'");
});

test('comments preserve continued Unicode, authors, rich text, anchor and visible state',async()=>{
  const {out,zip,xml}=await convert(commentRecords());assert.equal(out.losses.filter(l=>l.notify).length,0);
  assert.match(await zip.get('[Content_Types].xml').text(),/<Default Extension="vml" ContentType="application\/vnd.openxmlformats-officedocument.vmlDrawing"\/>/);
  assert.ok(zip.has('xl/comments1.xml'));const text=await zip.get('xl/comments1.xml').text(),vml=await zip.get('xl/drawings/vmlDrawing1.vml').text();
  assert.match(text,/<author>Author<\/author>/);assert.match(text,/<comment ref="B3" authorId="0"><text><r><rPr>/);assert.match(text,/>Hi文<\/t>/);
  assert.match(vml,/id="_x0000_s2049"/);assert.match(vml,/<x:Anchor>2, 16, 1, 4, 4, 32, 5, 9<\/x:Anchor>/);assert.match(vml,/visibility:visible/);assert.match(vml,/<x:Row>2<\/x:Row><x:Column>1<\/x:Column><x:Visible\/>/);assert.match(xml,/<legacyDrawing r:id=/);
});

test('a damaged comment loses that comment and leaves cells and hyperlinks readable',async()=>{
  const link=rec(0x1b8,cat(range(0,0),Buffer.alloc(16),dword(2),dword(8),dword(8),Buffer.from('Data!A1\0','utf16le')));
  const {out,xml,zip}=await convert([...commentRecords({badText:true}),rec(0x204,cat(words(0,0,0),str('kept'))),link]);
  assert.ok(!zip.has('xl/comments1.xml'));assert.match(xml,/location="Data!A1"/);assert.match(await zip.get('xl/sharedStrings.xml').text(),/>kept<\/t>/);
  assert.ok(out.losses.some(l=>l.id==='legacy-comments'&&l.notify&&l.place==='Data!B3'));
});

test('URL, file and internal hyperlinks keep relationship targets and fragment locations',async()=>{
  const target='https://example.test/a b\0',url=rec(0x1b8,cat(range(0,0),Buffer.alloc(16),dword(2),dword(1),dword(0x79eac9e0),Buffer.alloc(12),dword(target.length*2),Buffer.from(target,'utf16le')));
  const file=rec(0x1b8,cat(range(1,0),Buffer.alloc(16),dword(2),dword(1),dword(0x303),Buffer.alloc(12),words(2),dword(8),Buffer.from('foo.xls\0'),Buffer.alloc(24),dword(0)));
  const internal=rec(0x1b8,cat(range(2,0),Buffer.alloc(16),dword(2),dword(8),dword(8),Buffer.from('Data!B1\0','utf16le')));
  const {xml,zip,out}=await convert([url,file,internal]);assert.match(xml,/location="Data!B1"/);
  const rels=await zip.get('xl/worksheets/_rels/sheet1.xml.rels').text();assert.match(rels,/https:\/\/example.test\/a%20b/);assert.match(rels,/Target="\.\.\/\.\.\/foo.xls"/);assert.equal(out.losses.length,0);
});

test('validation lists translate NUL separators and relative conditions use the first range',async()=>{
  const list=cat([0x17,5,0],Buffer.from('A\0B\0C')),formula=cat(ref(0,0xc000),int(3),[13]);
  const {xml,out}=await convert([dv(3|128|256|0xc0000,list),dv(7,formula)]);
  assert.match(xml,/<formula1>&quot;A,B,C&quot;<\/formula1>/);assert.match(xml,/<formula1>B3&gt;3<\/formula1>/);assert.match(xml,/promptTitle="Pick"/);assert.match(xml,/sqref="B3:B5"/);assert.equal(out.losses.length,0);
});

test('classic CF keeps differential font/border/fill and CFEx rule priority and stopping',async()=>{
  const f=Buffer.alloc(118);f.writeInt32LE(-1,64);f.writeUInt32LE(2,68);f.writeUInt16LE(700,72);f.writeInt32LE(10,80);f.writeUInt32LE(0x98,88);f.writeUInt32LE(1,92);f.writeUInt32LE(1,96);
  const flags=(0x34000000|0x3fffff)&~(0x400|0x10000|0x20000|0x40000),dxf=cat(dword(flags),words(0),f,dword(1|(10<<16)),dword(0),words(1<<10,11|(12<<7)));
  const {xml,styles,out}=await convert([cfHeader(),cf(int(10),{dxf}),cf(int(20)),cfEx(0,2,3),cfEx(1,1,1)]);
  const rules=xml.match(/<cfRule[^>]*>/g);for(const attr of ['stopIfTrue="1"','priority="2"','type="cellIs"','operator="greaterThan"','dxfId="0"'])assert.ok(rules[0].includes(attr));assert.ok(rules[1].includes('priority="1"'));assert.ok(!rules[1].includes('stopIfTrue'));
  assert.match(styles,/<dxfs count="1">/);assert.match(styles,/<b val="1"\/>/);assert.match(styles,/<left style="thin"><color indexed="10"\/><\/left>/);assert.match(styles,/<patternFill patternType="solid"><fgColor indexed="11"\/><bgColor indexed="12"\/>/);assert.equal(out.losses.length,0);
});

test('CF12 color scales, data bars and icon sets retain thresholds and options',async()=>{
  const header=rec(0x879,cat(Buffer.alloc(12),words(3,2),range(0,0,10,0),words(1),range(0,0,10,0)));
  const bar=cat(words(42),[0,0,10,90],color('63c384'),threshold(2),threshold(3));
  const scale=cat(words(90),[0,2,2,3],threshold(2),f64(0),threshold(3),f64(1),f64(0),color('ff0000'),f64(1),color('00ff00'));
  const icon=cat(words(52),[0,3,0,5],...[[4,0],[4,33],[4,67]].map(([t,v])=>cat(threshold(t,v),[1],dword(0))));
  const {xml,out}=await convert([header,cf12(4,bar,3),cf12(3,scale,2),cf12(6,icon,1)]);
  assert.match(xml,/<dataBar minLength="10" maxLength="90" showValue="0"><cfvo type="min"\/><cfvo type="max"\/><color rgb="FF63C384"\/>/);
  assert.match(xml,/<colorScale><cfvo type="min"\/><cfvo type="max"\/><color rgb="FFFF0000"\/><color rgb="FF00FF00"\/>/);
  assert.match(xml,/<iconSet iconSet="3Arrows" showValue="0" reverse="1"><cfvo type="percent" val="0"\/>/);assert.equal(out.losses.length,0);
});

test('bad validation and CF records are local losses with other rules retained',async()=>{
  const {xml,out}=await convert([dv(1,int(1),int(10)),rec(0x1be,u(1)),cfHeader(),cf(int(2)),rec(0x1b1,u(2))]);
  assert.match(xml,/<dataValidations count="1">/);assert.match(xml,/<formula>2<\/formula>/);assert.equal(out.losses.filter(l=>l.notify).length,2);
});

test('print setup preserves margins, orientation, headers, page breaks and fit-to-page',async()=>{
  const setup=cat(words(9,125,3,1,0,0x2a9,300,600),f64(.25),f64(.35),words(2));
  const {xml,out}=await convert([rec(0x81,words(0x101)),rec(0xa1,setup),rec(0x26,f64(.4)),rec(0x14,str('&LReport&CPage &P')),rec(0x15,str('&R&D')),rec(0x1a,cat(words(1,3,0,65535))),rec(0x1b,cat(words(1,5,0,255)))]);
  assert.match(xml,/<pageSetUpPr autoPageBreaks="1" fitToPage="1"\/>/);assert.match(xml,/<pageMargins left="0.4"/);assert.match(xml,/orientation="landscape"/);assert.match(xml,/cellComments="atEnd"/);assert.match(xml,/useFirstPageNumber="1"/);assert.match(xml,/<oddHeader>&amp;LReport&amp;CPage &amp;P<\/oddHeader>/);assert.match(xml,/<rowBreaks[^>]*><brk id="5" min="0" max="255" man="1"\/>/);assert.match(xml,/<colBreaks[^>]*><brk id="3" min="0" max="65535" man="1"\/>/);assert.ok(xml.indexOf('<rowBreaks')<xml.indexOf('<colBreaks'));assert.equal(out.losses.length,0);
});

test('fNoPls ignores undefined printer values, and protection keeps hashes and allowed actions',async()=>{
  const setup=cat(words(65535,65535,0,1,1,4,0,0),f64(.5),f64(.5),words(0));
  const feat=rec(0x867,cat(Buffer.alloc(12),words(2),[1],dword(0xffffffff),dword(0x5404)));
  const {xml,book,out}=await convert([rec(0xa1,setup),rec(0x12,words(1)),rec(0x13,words(0x83af)),feat],[rec(0x12,words(1)),rec(0x19,words(1)),rec(0x13,words(0xabcd))]);
  assert.doesNotMatch(xml,/paperSize=|scale=|orientation=/);assert.match(xml,/<sheetProtection sheet="1" password="83AF"/);assert.match(xml,/formatCells="0"/);assert.match(xml,/autoFilter="0"/);assert.match(xml,/selectLockedCells="0"/);
  assert.match(book,/<workbookProtection lockStructure="1" lockWindows="1" workbookPassword="ABCD"\/>/);assert.equal(out.losses.length,0);
});

test('AutoFilter uses its built-in range name and keeps both comparison operands',async()=>{
  const tokens=cat([0x3b],words(0,1,9,0,2)),name=Buffer.alloc(14);name.writeUInt16LE(0x21);name[3]=1;name.writeUInt16LE(tokens.length,4);name.writeUInt16LE(1,8);
  const globals=[rec(0x1ae,words(1,0x401)),rec(0x17,words(1,0,0,0)),rec(0x18,cat(name,[0,13],tokens))];
  const filter=rec(0x9e,cat(words(1,0),[4,4],f64(10),[4,1],f64(20)));
  const {xml,out}=await convert([rec(0x9d,words(3)),filter],globals);assert.match(xml,/<autoFilter ref="A2:C10"><filterColumn colId="1"><customFilters and="1"><customFilter operator="greaterThan" val="10"\/><customFilter operator="lessThan" val="20"\/>/);assert.equal(out.losses.length,0);
});

test('font definitions and rich runs follow Office font child order',async()=>{
  const font=cat(words(200,2,10,700,1),[1,2,0,0,5,0],Buffer.from('Arial'));
  const {styles,zip}=await convert(commentRecords(),[rec(0x31,font)]);
  assert.match(styles,/<font><b\/><i\/><u val="single"\/><vertAlign val="superscript"\/><sz val="10"\/><color indexed="10"\/><name val="Arial"\/><family val="2"\/><\/font>/);
  assert.match(await zip.get('xl/comments1.xml').text(),/<rPr><b\/><i\/><u val="single"\/><vertAlign val="superscript"\/><sz val="10"\/><color indexed="10"\/><rFont val="Arial"\/>/);
});


test('CF12 flags use MS-XLS data-bar bits and distinguish signs from framed traffic lights',async()=>{
  const header=rec(0x879,cat(Buffer.alloc(12),words(3,2),range(0,0,10,0),words(1),range(0,0,10,0)));
  const bar=flags=>cat(words(42),[0,flags,10,90],color('63c384'),threshold(2),threshold(3));
  const icon=id=>cat(words(52),[0,3,id,0],...[[4,0],[4,33],[4,67]].map(([t,v])=>cat(threshold(t,v),[1],dword(0))));
  const {xml,out}=await convert([header,cf12(4,bar(2)),cf12(6,icon(4)),cf12(6,icon(5)),cf12(4,bar(1))]);
  assert.match(xml,/<dataBar minLength="10" maxLength="90">/);
  assert.match(xml,/<iconSet iconSet="3Signs">/);assert.match(xml,/<iconSet iconSet="3TrafficLights2">/);
  assert.equal(out.losses.filter(l=>l.notify).length,1);assert.match(out.losses.find(l=>l.notify).detail.message,/Right-to-left/);
});


test('damaged view and format metadata stays within Excel limits without losing cells',async()=>{
  const sel=(pane,row)=>rec(0x1d,cat([pane],words(row,0,0,1,row,row),[0,0]));
  const {xml,styles}=await convert([rec(0x7d,words(0,0,65535,0,0)),rec(0x41,words(1,1,1,1,0)),sel(3,0),sel(2,1),sel(1,2),sel(0,3),sel(3,4)],[rec(0x293,cat(words(0x8000),[255,0]))]);
  assert.match(xml,/width="255"/);assert.equal((xml.match(/<selection /g)||[]).length,4);assert.match(xml,/<selection pane="topLeft" activeCell="A5"/);
  assert.doesNotMatch(styles,/builtinId="255"/);assert.match(styles,/name="Style 255"/);
});

test('external links omit empty sheet lists and isolate malformed sheet names',async()=>{
  const globals=[rec(0x1ae,cat(words(0),str('empty.xls'))),rec(0x1ae,cat(words(2),str('other.xls'),str('A'.repeat(32)),str('Good'))),rec(0x17,words(1,1,1,1))];
  const {zip,out}=await convert([],globals),empty=await zip.get('xl/externalLinks/externalLink1.xml').text(),other=await zip.get('xl/externalLinks/externalLink2.xml').text();
  assert.doesNotMatch(empty,/<sheetNames/);assert.match(other,/<sheetNames><sheetName val="Good"\/><\/sheetNames>/);assert.doesNotMatch(other,/A{32}/);assert.equal(out.losses.length,0);
});

test('external workbook paths escape literal percent signs and retain existing URL escapes',async()=>{
  const {zip}=await convert([],[rec(0x1ae,cat(words(1),str('Costs% (2020)%20.xls'),str('Data')))]);
  const rels=await zip.get('xl/externalLinks/_rels/externalLink1.xml.rels').text();
  assert.match(rels,/Target="Costs%25%20\(2020\)%20.xls"/);
});

test('hyperlink frame metadata is consumed without writing an invalid SpreadsheetML attribute',async()=>{
  const link=rec(0x1b8,cat(range(0,0),Buffer.alloc(16),dword(2),dword(0x88),dword(6),Buffer.from('_self\0','utf16le'),dword(8),Buffer.from('Data!A1\0','utf16le')));
  const {xml,out}=await convert([link]);assert.match(xml,/location="Data!A1"/);assert.doesNotMatch(xml,/targetFrame/);assert.equal(out.losses.length,0);
});


test('Office XFExt light theme slots preserve their cached white fill instead of turning it black',()=>{
  // The real MixedGasReport XF caches palette 9 (white) and extends it with theme 0.
  const xf=Buffer.alloc(20);xf.writeUInt32LE(1<<26,14);xf.writeUInt16LE(9,18);
  const head=Buffer.alloc(20);head.writeUInt16LE(1,18);
  const extension=cat(words(4,20),words(3,0),dword(0),Buffer.alloc(8));
  const book=X.read(workbook([rec(0xe0,xf),rec(0x87d,cat(head,extension))],[]));book.theme=true;
  assert.match(X.styles(book,[]).xml,/<patternFill patternType="solid"><fgColor theme="0"/);
  assert.deepEqual([0,1,2,3].map(n=>X.colorValue(3,n,true).theme),[0,1,2,3]);
});


test('unconverted header pictures are a content loss, while escaped picture tokens are text',async()=>{
  const first=await convert([rec(0x14,str('&L&G'))]),second=await convert([rec(0x14,str('&&G'))]);
  assert.ok(first.out.losses.some(l=>l.id==='legacy-header-footer-pictures'&&l.notify));assert.equal(second.out.losses.length,0);
});
