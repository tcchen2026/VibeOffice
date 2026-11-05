// Isolate C1's carried anchor/gradient corrections in DrawingML and Word VML.
import fs from 'node:fs';import path from 'node:path';import {createHash} from 'node:crypto';import {spawnSync} from 'node:child_process';import {load} from './load.mjs';
const B=load(),out=process.argv[2]||'tmp/legacy/artifacts/C1-office-shapes',pairs=[],hashes=[];fs.mkdirSync(out,{recursive:true});
const shapes=[4,5,6,7].map((type,i)=>({id:1025+i,type:1,flags:0,properties:new Map([[384,{value:type}],[385,{value:0x2020ff}],[387,{value:0xffe0a0}],[395,{value:-45*65536}],[397,{value:32768}],[398,{value:32768}],[399,{value:32768}],[400,{value:32768}],[4,{value:i===0?90*65536:0}]])}));
for(const kind of ['ppt','doc']){
  const source='tmp/legacy/artifacts/B-office/'+kind+'-control.'+kind+'x',original=fs.readFileSync(source),zip=await L.zip.read(original),files=new Map();for(const [name,e]of zip)files.set(name,await e.bytes());
  if(kind==='ppt'){
    const head='<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';
    const body=shapes.map((s,i)=>B.el('p:sp',{},B.el('p:nvSpPr',{},B.el('p:cNvPr',{id:i+2,name:'Fill type '+[4,5,6,7][i]})+'<p:cNvSpPr/><p:nvPr/>')+B.el('p:spPr',{},B.shapes.dml(s,{x:700000+(i%2)*4200000,y:500000+Math.floor(i/2)*2600000,cx:3200000,cy:1500000})))).join('');
    const xml=L.zip.decode(files.get('ppt/slides/slide1.xml'));files.set('ppt/slides/slide1.xml',L.zip.encode(xml.replace(/<p:spTree\b[^>]*>[\s\S]*?<\/p:spTree>/,B.el('p:spTree',{},head+body))));
  }else{
    const body=shapes.map((s,i)=>'<w:p><w:r><w:pict xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">'+B.shapes.vml(s,{x:(i%2)*3000000,y:Math.floor(i/2)*2000000,cx:2600000,cy:1200000})+'</w:pict></w:r></w:p>').join('');
    const xml=L.zip.decode(files.get('word/document.xml'));files.set('word/document.xml',L.zip.encode(xml.replace(/<w:body>[\s\S]*?<w:sectPr/, '<w:body>'+body+'<w:sectPr')));
  }
  const blob=await L.zip.write([...files].map(([name,data])=>({name,data}))),bytes=new Uint8Array(await blob.arrayBuffer()),file=path.resolve(out,kind+'-rotation-gradients.'+kind+'x');fs.writeFileSync(file,bytes);
  const check=spawnSync('python3',['-I','tools/ooxml/package.py','check',file],{encoding:'utf8'});fs.writeFileSync(file+'.check.txt',check.stdout+check.stderr);if(check.status!==0)throw Error('Invalid diagnostic');pairs.push({name:kind+'-rotation-gradients',file});hashes.push({file,sha256:createHash('sha256').update(bytes).digest('hex'),source,sourceHash:createHash('sha256').update(original).digest('hex')});
}
fs.writeFileSync(path.join(out,'pairs.json'),JSON.stringify(pairs,null,2));fs.writeFileSync(path.join(out,'hashes.json'),JSON.stringify(hashes,null,2));fs.writeFileSync(path.join(out,'CHECKLIST.md'),'# I24.C1 shared shape corrections\n\nTwo diagnostic files, OFFICE_MAX=2.\n\n- [ ] Word and PowerPoint open without repair.\n- [ ] First shape rotates 90 degrees while retaining the binary anchor bounds.\n- [ ] Fill kinds 4 and 7 are linear; 5 and 6 follow bounds/outline.\n- [ ] Record job IDs and hashes; inspect PDFs.\n');
