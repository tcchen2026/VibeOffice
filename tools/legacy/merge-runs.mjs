// Overlay targeted checks on a full run without copying large package artifacts.
import fs from 'node:fs';
import path from 'node:path';
const [destination,...inputs]=process.argv.slice(2),rows=new Map(),provenance=[];
for(const input of inputs){const dir=path.resolve(input),run=fs.readFileSync(path.join(dir,'runs.jsonl'),'utf8').trim().split('\n').map(JSON.parse);provenance.push({directory:dir,cases:run.length});
  for(const row of run){if(row.status==='converted')row.pkg||=path.join(dir,'packages',row.sha256+(row.mutation?'-'+row.mutation:'')+'.'+row.type);row.run=dir;rows.set(row.file+'|'+(row.mutation||''),row);}
}
fs.mkdirSync(destination,{recursive:true});fs.writeFileSync(path.join(destination,'runs.jsonl'),[...rows.values()].map(r=>JSON.stringify(r)).join('\n')+'\n');
fs.writeFileSync(path.join(destination,'provenance.json'),JSON.stringify(provenance,null,2)+'\n');
console.log(JSON.stringify({cases:rows.size,inputs:provenance,status:[...rows.values()].reduce((a,r)=>{a[r.status]=(a[r.status]||0)+1;return a;},{})}));
