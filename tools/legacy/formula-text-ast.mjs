// Independent Ledger parse of changed text. Parser limits are reported, not hidden.
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
vm.runInThisContext(fs.readFileSync('public/ledger/js/formula.js','utf8'));
const F=L.formula,dir=process.argv[2],pairs=JSON.parse(fs.readFileSync(path.join(dir,'formula-pairs.json'),'utf8'));
function normal(x,booleans=false){
  if(Array.isArray(x))return x.map(v=>normal(v,booleans));
  if(!x||typeof x!=='object')return x;
  if(x.t==='paren')return normal(x.a,booleans);
  if(booleans&&x.t==='fn'&&!x.args.length&&['TRUE','FALSE'].includes(x.name))return {t:'bool',v:x.name==='TRUE'};
  const out=Object.fromEntries(Object.entries(x).filter(([k])=>k!=='raw').map(([k,v])=>[k,normal(v,booleans)]));
  if(out.t==='union')out.items=out.items.flatMap(v=>v.t==='union'?v.items:[v]);
  return out;
}
const report={same:0,wrappedUnionSame:0,parserLimits:[],different:[],libreoffice:{sameAfterBooleanAndQuoteNormalization:0,limits:[],different:[]}};
for(const p of pairs){
  let a,b;
  try{a=normal(F.parse(p.before));b=normal(F.parse(p.after));}
  catch(e){
    // Bare comma unions are valid Excel names; Ledger's parser currently expects
    // an enclosing pair. Wrapping the whole expression tests its operand tree.
    try{a=normal(F.parse('('+p.before+')'));b=normal(F.parse('('+p.after+')'));}
    catch{report.parserLimits.push({...p,error:e.message});continue;}
    if(JSON.stringify(a)===JSON.stringify(b)){report.wrappedUnionSame++;continue;}
  }
  if(JSON.stringify(a)===JSON.stringify(b))report.same++;else report.different.push({...p,a,b});
}
for(const p of JSON.parse(fs.readFileSync(path.join(dir,'libreoffice-pairs.json'),'utf8'))){
  try{const a=normal(F.parse(p.converted),true),b=normal(F.parse(p.libreoffice),true);if(JSON.stringify(a)===JSON.stringify(b))report.libreoffice.sameAfterBooleanAndQuoteNormalization++;else report.libreoffice.different.push({...p,a,b});}
  catch(e){report.libreoffice.limits.push({...p,error:e.message});}
}
fs.writeFileSync(path.join(dir,'asts.json'),JSON.stringify(report,null,2)+'\n');
console.log(JSON.stringify({...report,parserLimits:report.parserLimits.length,different:report.different.length,libreoffice:{...report.libreoffice,limits:report.libreoffice.limits.length,different:report.libreoffice.different.length}}));
if(report.different.length)process.exitCode=1;
