// Match the current corpus outputs to the exact Office-accepted uncompressed parts.
import fs from 'node:fs';
import path from 'node:path';
import {createHash} from 'node:crypto';
import {load} from './load.mjs';
load();const base=path.resolve('tmp/legacy/artifacts'),rows=fs.readFileSync(path.join(base,'C2-final/runs.jsonl'),'utf8').trim().split('\n').map(JSON.parse),proof=[];
const hash=b=>createHash('sha256').update(b).digest('hex');
async function parts(file){const zip=await L.zip.read(fs.readFileSync(file)),result={};for(const [name,entry]of zip)result[name]=hash(await entry.bytes());return result;}
for(const round of ['C2-office-final','C2-office-final-main']){
  const dir=path.join(base,round),diff=JSON.parse(fs.readFileSync(path.join(dir,'diff.json'),'utf8')),pairs=JSON.parse(fs.readFileSync(path.join(dir,'pairs.json'),'utf8'));
  for(const input of diff){
    const source=rows.find(r=>!r.mutation&&r.file===input.source),pair=pairs.find(p=>p.file===input.file||p.file===input.file.replace(/\.xlsx$/,'-print-control.xlsx'));let accepted=pair.file,resultFile=path.join(dir,'results',pair.name,'file/result.json'),diagnostic=!source.stats.cells;
    if(input.source.endsWith('__user_defined_function.xls')){accepted=path.join(base,'C2-office-bisect-2/user-function-only-B51.xlsx');resultFile=path.join(base,'C2-office-bisect-2/results/user-function-only-B51/file/result.json');}
    const a=await parts(source.pkg),b=await parts(diagnostic?input.file:accepted),differences=[...new Set([...Object.keys(a),...Object.keys(b)])].filter(n=>a[n]!==b[n]);
    const result=JSON.parse(fs.readFileSync(resultFile,'utf8')),pdf=path.join(path.dirname(resultFile),'office.pdf');
    if(differences.length||!result.canOpenWithoutRepair||result.pdfStatus!=='ok')throw Error('Office proof differs: '+input.source+' '+differences.join(','));
    proof.push({source:input.source,currentPackage:source.pkg,currentHash:hash(fs.readFileSync(source.pkg)),acceptedPackage:accepted,acceptedHash:hash(fs.readFileSync(accepted)),job:result.id,result:resultFile,pdf,pdfHash:hash(fs.readFileSync(pdf)),diagnostic,partDifferences:differences,notes:diagnostic?'Source has no cells; separate print control adds only A1. Source original opened in round 1 but could not export a PDF.':null});
  }
}
fs.writeFileSync(path.join(base,'C2-final/office-proof.json'),JSON.stringify(proof,null,2)+'\n');console.log(JSON.stringify({inputs:proof.length,converted:proof.filter(r=>!r.diagnostic).length,printControls:proof.filter(r=>r.diagnostic).length,allMatch:true}));
