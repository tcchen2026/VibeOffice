// In-file slide order/text/geometry, independent of the OOXML writer.
import fs from 'node:fs';import path from 'node:path';import {load}from './load.mjs';
const B=load(),manifest=JSON.parse(fs.readFileSync('/work/corpora/legacy-manifest.json','utf8')),out=[];
for(const entry of manifest.filter(x=>x.format==='powerpoint')){
 const row={file:entry.file,sha256:entry.sha256,size:entry.size};out.push(row);
 try{
  const bytes=fs.readFileSync('/work/corpora/'+entry.file),sniff=B.sniff(bytes);row.sniff=sniff;
  if(sniff.format!=='ppt'||sniff.version!=='PowerPoint97+'||sniff.encrypted)continue;
  const book=B.ppt.structure(B.cfb(bytes),[]);
  Object.assign(row,{size:entry.size,dimensions:book.size,notesDimensions:book.notesSize,edits:book.edits.length,masters:book.masters.length,notes:book.notes.length,fonts:[...book.fonts.values()],issues:book.losses,slides:book.slides.map(s=>({id:s.id,ref:s.ref,master:s.masterId,notes:s.notesId,hidden:s.hidden,text:s.text,scheme:s.scheme}))});
 }catch(e){if(!['LegacyError','CFBError'].includes(e.name))throw e;row.error={code:e.code,message:e.message};}
}
const file=process.argv[2];fs.mkdirSync(path.dirname(file),{recursive:true});fs.writeFileSync(file,JSON.stringify(out,null,2));console.log(JSON.stringify({files:out.length,read:out.filter(r=>r.slides).length,slides:out.reduce((n,r)=>n+(r.slides?.length||0),0)}));
