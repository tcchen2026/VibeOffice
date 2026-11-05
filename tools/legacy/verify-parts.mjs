// Diagnostic packages: isolate shared metadata and geometry from the unfinished converters.
import fs from 'node:fs';
import path from 'node:path';
import {createHash} from 'node:crypto';
import {spawnSync} from 'node:child_process';
import {load} from './load.mjs';
const B=load(), L=globalThis.L, out=process.argv[2]||'tmp/legacy/artifacts/B-office';fs.mkdirSync(out,{recursive:true});
const rows=fs.readFileSync('tmp/legacy/artifacts/B/runs.jsonl','utf8').trim().split('\n').map(JSON.parse).filter(r=>!r.mutation&&r.shared?.propertyParts===3);
const bases={doc:'word/libreoffice__4d559ed1b2aa__customshape-position.docx',xls:'excel/poi__35d1f8273639__customIndexedColors.xlsx',ppt:'powerpoint/libreoffice__c385a5f717e5__picture-placeholder-custom-prompt.pptx'};
const pairs=[], hashes=[],diagnostics=[];
const hash=b=>createHash('sha256').update(b).digest('hex');
for(const [kind,base]of Object.entries(bases)){
  const source=rows.filter(r=>r.sniff?.format===kind&&!r.sniff?.encrypted).sort((a,b)=>a.size-b.size)[0];
  const cfb=B.cfb(new Uint8Array(fs.readFileSync('/work/corpora/'+source.file))),props=B.propertyParts(cfb),original=fs.readFileSync('/work/corpora/'+base),zip=await L.zip.read(original);
  const files=new Map();for(const [name,e]of zip)files.set(name,await e.bytes());
  for(const p of props.parts)files.set(p.name,L.zip.encode(p.data));
  const propertyNames=new Set(props.parts.map(p=>'/'+p.name));
  let types=L.zip.decode(files.get('[Content_Types].xml')).replace(/<Override\b[^>]*\/>/g,s=>[...propertyNames].some(n=>s.includes('PartName="'+n+'"'))?'':s);
  types=types.replace('</Types>',props.parts.map(p=>B.el('Override',{PartName:'/'+p.name,ContentType:p.type})).join('')+'</Types>');files.set('[Content_Types].xml',L.zip.encode(types));
  let rels=L.zip.decode(files.get('_rels/.rels')).replace(/<Relationship\b[^>]*\/>/g,s=>/Type="[^"]*(?:core-properties|extended-properties|custom-properties)"/.test(s)?'':s);
  const ids=new Set([...rels.matchAll(/Id="([^"]+)"/g)].map(m=>m[1]));let next=1;
  for(const p of props.parts){while(ids.has('rId'+next))next++;const id='rId'+next++;ids.add(id);rels=rels.replace('</Relationships>',B.el('Relationship',{Id:id,Type:p.relation.includes('://')?p.relation:B.ns.officeRel+p.relation,Target:p.name})+'</Relationships>');}
  files.set('_rels/.rels',L.zip.encode(rels));
  const save=async(label,entries)=>{const file=path.resolve(out,label+'.'+kind+'x'),blob=await L.zip.write([...entries].map(([name,data])=>({name,data}))),data=new Uint8Array(await blob.arrayBuffer());fs.writeFileSync(file,data);hashes.push({file,sha256:hash(data),binarySource:source.file,binaryHash:source.sha256});const check=spawnSync('python3',['-I','tools/ooxml/package.py','check',file],{encoding:'utf8'});fs.writeFileSync(file+'.check.txt',check.stdout+check.stderr);if(check.status!==0)throw Error('Package check failed: '+file);return file;};
  const file=await save(kind+'-metadata',files),control=path.resolve(out,kind+'-control.'+kind+'x');fs.writeFileSync(control,original);hashes.push({file:control,sha256:hash(original),source:base});pairs.push({name:kind+'-metadata',original:control,saved:file});
  if(kind==='ppt'){
    const el=B.el,shapes=[{type:13,properties:new Map([[327,{value:14000}],[328,{value:6000}],[385,{value:0x8080ff}]]),flags:0},{type:19,properties:new Map([[327,{value:90*65536}],[328,{value:10*65536}]]),flags:0},{type:60,properties:new Map(),flags:64}];
    const body='<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'+shapes.map((s,i)=>el('p:sp',{},el('p:nvSpPr',{},el('p:cNvPr',{id:i+2,name:'Shape '+(i+1)})+'<p:cNvSpPr/><p:nvPr/>')+el('p:spPr',{},B.shapes.dml(s,{x:300000+i*2800000,y:700000,cx:2300000,cy:2100000})))).join('');
    const slide=L.zip.decode(files.get('ppt/slides/slide1.xml'));files.set('ppt/slides/slide1.xml',L.zip.encode(slide.replace(/<p:spTree\b[^>]*>[\s\S]*?<\/p:spTree>/,el('p:spTree',{},body))));
    diagnostics.push({name:'ppt-geometry',file:await save('ppt-geometry',files)});
  }
  if(kind==='doc'){
    const shape={id:1025,type:13,properties:new Map([[327,{value:14000}],[328,{value:6000}]]),flags:0},vml=B.shapes.vml(shape,{x:0,y:0,cx:2540000,cy:1270000});
    const paragraph='<w:p><w:r><w:pict xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">'+vml+'</w:pict></w:r></w:p>';
    const doc=L.zip.decode(files.get('word/document.xml'));files.set('word/document.xml',L.zip.encode(doc.replace('<w:body>','<w:body>'+paragraph)));
    diagnostics.push({name:'doc-geometry',file:await save('doc-geometry',files)});
  }
}
fs.writeFileSync(path.join(out,'pairs.json'),JSON.stringify([...pairs,...diagnostics],null,2)+'\n');fs.writeFileSync(path.join(out,'hashes.json'),JSON.stringify(hashes,null,2)+'\n');
fs.writeFileSync(path.join(out,'CHECKLIST.md'),'# I24.B shared-parts diagnostic round\n\nEight files, OFFICE_MAX=2. These are NOT converter successes.\n\n- [ ] Three existing Office-authored OOXML controls open without repair.\n- [ ] Three metadata replacements open without repair; only docProps, root metadata relationships and content types differ.\n- [ ] Word VML native adjusted arrow opens without repair and appears.\n- [ ] PPT slide1 three generated shapes (adjusted arrow, arc, star) opens without repair and appears.\n- [ ] Record all jobs, hashes and PDFs.\n\nBinary metadata sources and exact SHA256s: hashes.json. Output parts checked with tools/ooxml/package.py.\n');
console.log(JSON.stringify({out,files:hashes.length,pairs: pairs.length,diagnostics:diagnostics.length}));
