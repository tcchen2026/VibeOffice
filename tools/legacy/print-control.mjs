// Office's PDF-only gate cannot export an empty workbook. Add one clearly
// identified diagnostic cell to a separate copy; preserve the source package.
import fs from 'node:fs';
import path from 'node:path';
import {createHash} from 'node:crypto';
import {load} from './load.mjs';
load();const dir=process.argv[2],pairs=JSON.parse(fs.readFileSync(path.join(dir,'pairs.json'),'utf8')),rows=JSON.parse(fs.readFileSync(path.join(dir,'diff.json'),'utf8')),controls=[];
for(const row of rows){if(row.stats.cells)continue;const zip=await L.zip.read(fs.readFileSync(row.file)),parts=[];
  for(const [name,entry]of zip){let data=await entry.bytes();if(name==='xl/worksheets/sheet1.xml'){const xml=L.zip.decode(data),replacement='<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Name-only workbook: print control</t></is></c></row></sheetData>';if(!xml.includes('<sheetData></sheetData>'))throw Error('Expected empty worksheet');data=L.zip.encode(xml.replace('<sheetData></sheetData>',replacement));}parts.push({name,data});}
  const bytes=new Uint8Array(await (await L.zip.write(parts)).arrayBuffer()),file=row.file.replace(/\.xlsx$/,'-print-control.xlsx');fs.writeFileSync(file,bytes);
  const pair=pairs.find(p=>p.file===row.file);pair.file=file;pair.name+='-print-control';controls.push({source:row.file,sourceHash:row.outputHash,file,sha256:createHash('sha256').update(bytes).digest('hex'),change:'Only sheet1 A1 gets a diagnostic print marker; names and all other parts are unchanged.'});
}
fs.writeFileSync(path.join(dir,'pairs.json'),JSON.stringify(pairs,null,2)+'\n');fs.writeFileSync(path.join(dir,'print-controls.json'),JSON.stringify(controls,null,2)+'\n');
fs.appendFileSync(path.join(dir,'CHECKLIST.md'),'\nEmpty-workbook diagnostic controls (not original content):\n'+controls.map(c=>'- '+path.basename(c.file)+': '+c.change+' SHA256 '+c.sha256).join('\n')+'\n');console.log(JSON.stringify(controls));
