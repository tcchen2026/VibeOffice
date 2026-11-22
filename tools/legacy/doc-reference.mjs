// Read the pinned binary Word stories/counts independently of package writing.
// node tools/legacy/doc-reference.mjs OUTPUT.json
import fs from 'node:fs';
import path from 'node:path';
import {load} from './load.mjs';
const B=load(),corpus='/work/corpora',manifest=JSON.parse(fs.readFileSync(corpus+'/legacy-manifest.json','utf8')),rows=[];
for(const entry of manifest.filter(e=>e.format==='word')){
  const row={file:entry.file,sha256:entry.sha256,size:entry.size};
  try{
    const bytes=fs.readFileSync(path.join(corpus,entry.file));row.sniff=B.doc.sniff(bytes);
    if(row.sniff.format!=='doc'||row.sniff.version!=='Word97+'||row.sniff.encrypted){row.status='unavailable';rows.push(row);continue;}
    const result=B.inspect(bytes);row.status='read';row.features=result.features;row.references=result.references;
  }catch(e){row.status=['LegacyError','CFBError'].includes(e.name)?'refused':'crashed';row.code=e.code;row.message=e.message;}
  rows.push(row);
}
fs.writeFileSync(process.argv[2]||'tmp/legacy/artifacts/E1-source.json',JSON.stringify(rows));
const counts={};for(const r of rows)counts[r.status]=(counts[r.status]||0)+1;console.log(JSON.stringify({sources:rows.length,counts}));if(rows.some(r=>r.status==='crashed'))process.exitCode=1;
