// Select the exact source features affected by the post-census C2 corrections.
import fs from 'node:fs';
import {load} from './load.mjs';
const B=load(),run=process.argv[2]||'tmp/legacy/artifacts/C2',out=process.argv[3]||'tmp/legacy/C2-review-files.json';
const scopedOnly=process.argv.includes('--self-scoped-only'),parse=B.xls.formula;let scopedHit=false;
if(scopedOnly)B.xls.formula=(tokens,extra,context)=>parse(tokens,extra,{...context,externalName:(...args)=>{const result=context.externalName?.(...args),text=typeof result==='string'?result:result?.text;if(text?.startsWith("'")&&!text.startsWith("'["))scopedHit=true;return result;}});
const rows=fs.readFileSync(run+'/runs.jsonl','utf8').trim().split('\n').map(JSON.parse),selected=new Set(),details=[];
for(const row of rows){if(row.status!=='converted')continue;
  let bytes=new Uint8Array(fs.readFileSync('/work/corpora/'+row.file));
  if(row.mutation==='truncated')bytes=bytes.subarray(0,Math.floor(bytes.length*.57));
  if(row.mutation==='bitflip'){const seed=parseInt(row.sha256.slice(0,8),16);bytes[seed%bytes.length]^=1<<(seed%8);}
  if(row.mutation==='header')bytes.fill(255,44,76);
  const cfb=B.cfb(bytes),book=B.xls.read(cfb.stream('Workbook')||cfb.stream('Book')),reasons=[];
  if(scopedOnly){scopedHit=false;B.xls.prepareFormulas(book,[]);if(scopedHit)reasons.push('sheet-scoped self-book NameX');}
  for(const {id,r}of scopedOnly?[]:book.globals){
    if(id===0x18&&(!r.u16(4)||r.u16(0)&0xc))reasons.push('empty or macro NAME');
    if(id===0x23&&!(r.u16(0)&0x1a))reasons.push('external NAME');
    if(id===0xd&&r.u16(0)===2)reasons.push('CalcMode 2');
  }
  for(const {id,r}of scopedOnly?[]:book.sheets[0]?.records||[])if(id===0xd&&r.u16(0)===2)reasons.push('CalcMode 2');
  if(reasons.length){selected.add(row.file);details.push({file:row.file,mutation:row.mutation,reasons:[...new Set(reasons)]});}
}
fs.writeFileSync(out,JSON.stringify([...selected].sort(),null,2)+'\n');fs.writeFileSync(out+'.details.json',JSON.stringify(details,null,2)+'\n');
console.log(JSON.stringify({sources:selected.size,cases:details.length,reasons:details.reduce((a,r)=>{for(const k of r.reasons)a[k]=(a[k]||0)+1;return a;},{})}));
