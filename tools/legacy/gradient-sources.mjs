// Office-written binary gradient properties beside their embedded OOXML metro shape.
import fs from 'node:fs';import {load} from './load.mjs';import {auditShared} from './shared-audit.mjs';
const B=load(),rows=JSON.parse(fs.readFileSync('/work/corpora/legacy-manifest.json','utf8')),out='tmp/legacy/artifacts/C1-gradient-sources';fs.mkdirSync(out,{recursive:true});
const read=B.officeart.shapes,found=[];let source;
B.officeart.shapes=function(...args){const shapes=read(...args);const visit=items=>{for(const s of items){const type=s.properties.get(384)?.value,blob=s.properties.get(937)?.data;if([4,5,6,7].includes(type)&&blob&&found.length<20){const name=out+'/'+source.sha256+'-'+s.id+'.zip';fs.writeFileSync(name,blob);found.push({file:source.file,hash:source.sha256,id:s.id,type,properties:[...s.properties].filter(([id,p])=>id<448&&!p.data).map(([id,p])=>[id,p.value]),blob:name});}visit(s.children);}};visit(shapes);return shapes;};
for(const row of rows.filter(r=>['powerpoint','excel'].includes(r.format)).sort((a,b)=>(a.format==='powerpoint'?-1:1)-(b.format==='powerpoint'?-1:1))){source=row;try{const bytes=fs.readFileSync('/work/corpora/'+row.file);await auditShared(B,bytes,B.sniff(bytes));}catch{}if(found.length>=20)break;}
fs.writeFileSync(out+'/index.json',JSON.stringify(found,null,2));console.log('Embedded Office gradient examples:',found.length);
