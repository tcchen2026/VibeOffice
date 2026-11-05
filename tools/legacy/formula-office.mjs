// Prepare at most eight C2 Office inputs and diff each against the C1 writer.
// The baseline is one historical writer file, never a repository copy.
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import {execFileSync} from 'node:child_process';
import {createHash} from 'node:crypto';
import {load} from './load.mjs';
const [listing,outArg]=process.argv.slice(2),out=path.resolve(outArg),B=load(),rows=JSON.parse(fs.readFileSync(listing,'utf8'));
if(rows.length>8)throw Error('Office rounds contain at most eight files');
fs.mkdirSync(out,{recursive:true});
const current=fs.readFileSync('public/common/legacy/xls-write.js','utf8'),before=execFileSync('git',['show','1ac2d48:public/common/legacy/xls-write.js'],{encoding:'utf8'});
const hash=b=>createHash('sha256').update(b).digest('hex'),pairs=[],report=[];
async function parts(pkg){const zip=await L.zip.read(pkg),result={};for(const [name,entry]of zip)result[name]=hash(await entry.bytes());return result;}
for(const row of rows){
  const bytes=fs.readFileSync('/work/corpora/'+row.file),name=path.basename(row.file).replace(/\.[^.]+$/,'');
  vm.runInThisContext(before,{filename:'C1-xls-write.js'});const old=await B.xls.convert(bytes);
  vm.runInThisContext(current,{filename:'C2-xls-write.js'});const now=await B.xls.convert(bytes);
  const original=path.join(out,name+'-C1.'+old.type),file=path.join(out,name+'.'+now.type);
  fs.writeFileSync(original,old.pkg);fs.writeFileSync(file,now.pkg);
  const a=await parts(old.pkg),b=await parts(now.pkg),changes=[...new Set([...Object.keys(a),...Object.keys(b)])].filter(n=>a[n]!==b[n]).map(n=>({part:n,before:a[n]||null,after:b[n]||null}));
  report.push({source:row.file,sourceHash:hash(bytes),original,file,beforeHash:hash(old.pkg),outputHash:hash(now.pkg),changes,stats:now.stats,losses:now.losses});pairs.push({name,file});
}
fs.writeFileSync(path.join(out,'pairs.json'),JSON.stringify(pairs,null,2)+'\n');
fs.writeFileSync(path.join(out,'diff.json'),JSON.stringify(report,null,2)+'\n');
fs.writeFileSync(path.join(out,'CHECKLIST.md'),'# I24.C2 Office round\n\n'+
  '- Eight files maximum; OFFICE_MAX=2. No binary source is submitted.\n'+
  '- Before submission inspect diff.json: formula cells/caches, names, external links, calcPr and omitted false date1904 are expected changes. Other part changes need explanation.\n'+
  '- Each input must open without repair and produce a PDF; record result job IDs and input/PDF hashes.\n'+
  '- Inspect shared/array/table results, scoped names, sheet ranges and external cached results.\n\n'+
  report.map(r=>'- '+path.basename(r.file)+': '+r.outputHash).join('\n')+'\n');
console.log(JSON.stringify(report.map(r=>({file:r.source,changes:r.changes.map(c=>c.part),formulas:r.stats.formulas})),null,2));
