// Office diagnostic for the five native definitions with unused trailing commands.
import fs from 'node:fs';
import path from 'node:path';
import {createHash} from 'node:crypto';
import {spawnSync} from 'node:child_process';
import {load} from './load.mjs';
const B=load(), L=globalThis.L, out=process.argv[2]||'tmp/legacy/artifacts/B-native-office';
fs.mkdirSync(out,{recursive:true});
const source='tmp/legacy/artifacts/B-office/ppt-control.pptx', original=fs.readFileSync(source), zip=await L.zip.read(original),files=new Map();
for(const [name,e]of zip)files.set(name,await e.bytes());
const el=B.el,types=[42,45,48,51,104];
const body='<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'+types.map((type,i)=>{
  const shape={type,flags:0,properties:new Map([[327,{value:B.nativeGeometry[type].a[0]}],[385,{value:0x80c0ff}]])};
  return el('p:sp',{},el('p:nvSpPr',{},el('p:cNvPr',{id:i+2,name:'Native shape '+type})+'<p:cNvSpPr/><p:nvPr/>')+
    el('p:spPr',{},B.shapes.dml(shape,{x:900000+(i%3)*2800000,y:700000+Math.floor(i/3)*2800000,cx:1800000,cy:1500000})));
}).join('');
const slide=L.zip.decode(files.get('ppt/slides/slide1.xml'));
files.set('ppt/slides/slide1.xml',L.zip.encode(slide.replace(/<p:spTree\b[^>]*>[\s\S]*?<\/p:spTree>/,el('p:spTree',{},body))));
const file=path.resolve(out,'native-tails.pptx'),blob=await L.zip.write([...files].map(([name,data])=>({name,data}))),data=new Uint8Array(await blob.arrayBuffer());
fs.writeFileSync(file,data);
const check=spawnSync('python3',['-I','tools/ooxml/package.py','check',file],{encoding:'utf8'});
fs.writeFileSync(file+'.check.txt',check.stdout+check.stderr);if(check.status!==0)throw Error('Package check failed');
fs.writeFileSync(path.join(out,'pairs.json'),JSON.stringify([{name:'native-tails',file}],null,2)+'\n');
fs.writeFileSync(path.join(out,'hashes.json'),JSON.stringify({source,sourceHash:createHash('sha256').update(original).digest('hex'),file,sha256:createHash('sha256').update(data).digest('hex'),types},null,2)+'\n');
fs.writeFileSync(path.join(out,'CHECKLIST.md'),'# I24.B native geometry diagnostic\n\nOne file, OFFICE_MAX=2. No converter success claimed.\n\n- [ ] PowerPoint opens without repair and exports PDF.\n- [ ] Five shapes appear: callout2, accentCallout2, borderCallout2, accentBorderCallout2, curvedUpArrow.\n- [ ] Record result/job, input/PDF hashes and visual check.\n');
console.log(file);
