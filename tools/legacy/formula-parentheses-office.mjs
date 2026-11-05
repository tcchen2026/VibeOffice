// C2 renderer review: rerun changed accepted inputs and paired syntax controls.
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import {createHash} from 'node:crypto';
import {load} from './load.mjs';
import {words,rec,str,workbook,cfb} from '../../tests/legacy/xls-fixtures.mjs';
const B=load(),base=path.resolve('tmp/legacy/artifacts'),out=path.join(base,'C2-parens-office');
const qualified=process.argv.includes('--qualified-name');
const rows=folder=>new Map(fs.readFileSync(path.join(base,folder,'runs.jsonl'),'utf8').trim().split('\n').map(JSON.parse).filter(r=>!r.mutation).map(r=>[r.file,r]));
const previous=rows('C2-final'),current=rows('C2-parens-final'),proof=JSON.parse(fs.readFileSync(path.join(base,'C2-final/office-proof.json'),'utf8'));
const hash=b=>createHash('sha256').update(b).digest('hex');
async function parts(file){const z=await L.zip.read(fs.readFileSync(file)),r={};for(const [p,e]of z)r[p]=hash(await e.bytes());return r;}
const changed=[],unchanged=[];
for(const old of proof){
  const row=current.get(old.source),a=await parts(old.currentPackage),b=await parts(row.pkg),diff=[...new Set([...Object.keys(a),...Object.keys(b)])].filter(p=>a[p]!==b[p]);
  if(!diff.length){unchanged.push({...old,currentPackage:row.pkg,currentHash:hash(fs.readFileSync(row.pkg))});continue;}
  changed.push({name:path.basename(row.file,'.xls'),file:row.pkg,source:row.file,previous:old,changes:diff});
}
if(changed.length>8)throw Error('Split changed inputs into rounds of at most eight');
fs.mkdirSync(out,{recursive:true});
function round(name,items){
  const dir=path.join(out,name);fs.mkdirSync(dir,{recursive:true});
  fs.writeFileSync(path.join(dir,'pairs.json'),JSON.stringify(items.map(({name,file})=>({name,file})),null,2)+'\n');
  fs.writeFileSync(path.join(dir,'inputs.json'),JSON.stringify(items.map(v=>({...v,inputHash:hash(fs.readFileSync(v.file))})),null,2)+'\n');
  fs.writeFileSync(path.join(dir,'CHECKLIST.md'),'# I24.C2 formula-parentheses review\n\n- OFFICE_MAX=2; at most eight inputs.\n- Open without repair and export a PDF.\n- Only formula text changes in the real conversions; cells, caches, names, relationships and formatting stay intact.\n- Compare each PDF with the recorded previous accepted input or its paired before control.\n- Diagnostic covers precedence, associativity, explicit parentheses, percent/signs, unions, ranges/intersections and bare union names.\n\n'+items.map(v=>'- '+v.name+': '+hash(fs.readFileSync(v.file))).join('\n')+'\n');
}
if(!qualified)round('accepted',changed);
const source=[...current.keys()].find(n=>n.endsWith('__15228.xls')),extra=qualified?[]:[
  {name:'15228-before',file:previous.get(source).pkg,source,pair:'15228'},
  {name:'15228-after',file:current.get(source).pkg,source,pair:'15228'}
];
const cat=(...p)=>Buffer.concat(p.map(v=>Buffer.from(v))),n=v=>cat([0x1e],words(v)),ref=(r,c)=>cat([0x24],words(r,0xc000|c));
const fn=(id,count)=>cat([0x22,count],words(id)),area=(r1,r2,c1,c2)=>cat([0x25],words(r1,r2,0xc000|c1,0xc000|c2));
const cases=[
  ['precedence',cat(n(2),n(3),n(4),[5,3]),14],['lower child',cat(n(2),n(3),[3],n(4),[5]),20],
  ['right subtraction',cat(n(2),n(3),n(4),[4,4]),3],['right division',cat(n(20),n(3),n(2),[5,6]),20/6],
  ['left exponent',cat(n(2),n(3),[7],n(2),[7]),64],['right exponent',cat(n(2),n(3),n(2),[7,7]),512],
  ['negative base',cat(n(2),[0x13],n(2),[7]),4],['negative power',cat(n(2),n(2),[7,0x13]),-4],
  ['percent sign',cat(n(50),[0x14,0x13]),-.5],['sign percent',cat(n(50),[0x13,0x14]),-.5],
  ['double percent',cat(n(50),[0x14,0x14]),.005],['power percent',cat(n(2),n(3),[7,0x14]),.08],
  ['typed parentheses',cat(n(2),n(3),[3,0x15,0x15]),5],
  ['union argument',cat(ref(0,4),ref(1,4),[0x10],fn(4,1)),3],
  ['intersection',cat(area(0,1,4,5),area(0,1,5,5),[0x0f],fn(4,1)),12],
  ['unary reference',cat(ref(0,4),ref(0,4),[0x11,0x13]),-1],
  ['union name',cat([0x23],words(1,0),fn(4,1)),3],
  ['parenthesized refs',cat(ref(7,3),ref(6,4),[0x15,0x11],ref(8,4),[0x15],ref(7,5),[0x11,0x0f]),11],
  ['empty intersection',cat(ref(7,3),ref(6,4),[0x15,0x11],ref(9,4),[0x15],ref(8,5),[0x11,0x0f]),0,'e'],
  ['bare union',cat(ref(0,4),ref(1,4),[0x10]),15,'e']
];
const records=[rec(0x7d,words(0,0,7000,0,0))];
for(const [i,[label,tokens,value,type]]of cases.entries()){
  records.push(rec(0x204,cat(words(i,0,0),str(label))));const b=Buffer.alloc(22);b.writeUInt16LE(i);b.writeUInt16LE(1,2);b.writeUInt16LE(tokens.length,20);
  if(type==='e'){b.writeUInt16LE(65535,12);b[6]=2;b[8]=value;}else b.writeDoubleLE(value,6);records.push(rec(6,cat(b,tokens)));
}
for(const [r,c,v]of [[0,4,1],[1,4,2],[0,5,4],[1,5,8],[7,4,11]]){const b=Buffer.alloc(14);b.writeUInt16LE(r);b.writeUInt16LE(c,2);b.writeDoubleLE(v,6);records.push(rec(0x203,b));}
const ref3D=(r,c)=>cat([0x3a],words(0,r,c));
const name='Areas',tokens=qualified?cat(ref3D(0,4),ref3D(1,4),[0x10]):cat(ref(0,4),ref(1,4),[0x10]),header=Buffer.alloc(14);header[3]=name.length;header.writeUInt16LE(tokens.length,4);
const globals=qualified?[rec(0x1ae,words(1,0x401)),rec(0x17,words(1,0,0,0))]:[];
globals.push(rec(0x18,cat(header,[1],Buffer.from(name,'utf16le'),tokens)));
const binary=cfb(workbook(globals,[{name:'Precedence',records}]));
const before=fs.readFileSync('tmp/legacy/C2-before-parens/xls-formula.js','utf8'),after=fs.readFileSync('public/common/legacy/xls-formula.js','utf8');
for(const [version,code]of [['before',before],['after',after]]){
  vm.runInThisContext(code);const r=await B.xls.convert(binary);if(r.stats.formulas.converted!==cases.length)throw Error('Diagnostic formulas dropped');
  const file=path.join(out,'precedence-'+(qualified?'qualified-':'')+version+'.xlsx');fs.writeFileSync(file,r.pkg);extra.push({name:'precedence-'+version,file,pair:'precedence',stats:r.stats,losses:r.losses});
}
round(qualified?'qualified':'controls',extra);
if(!qualified){
  fs.writeFileSync(path.join(out,'unchanged-proof.json'),JSON.stringify(unchanged,null,2)+'\n');
  fs.writeFileSync(path.join(out,'diagnostic-cases.json'),JSON.stringify(cases.map(([label,,value,type])=>({label,value,type:type||'n'})),null,2)+'\n');
}
console.log(JSON.stringify({changed:changed.length,unchanged:unchanged.length,controls:extra.length,diagnosticFormulas:cases.length}));
