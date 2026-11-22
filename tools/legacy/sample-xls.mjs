// Named development samples, separate from the once-per-feature full corpus run.
import fs from 'node:fs';
import path from 'node:path';
import {createHash} from 'node:crypto';
import {spawnSync} from 'node:child_process';
import {load} from './load.mjs';
const B=load(),out=process.argv[2]||'tmp/legacy/artifacts/C1-samples';fs.mkdirSync(out,{recursive:true});
const sources=JSON.parse(fs.readFileSync(process.argv[3]||'tmp/legacy/C1-samples.json','utf8')),rows=[];
for(const item of sources){const start=performance.now(),bytes=new Uint8Array(fs.readFileSync('/work/corpora/'+item.file)),row={file:item.file,sourceHash:item.sha256};
  try{const result=await B.xls.convert(bytes,{});row.pkg=path.resolve(out,item.file.split('/').at(-1).replace(/\.xls$/,'.'+result.type));fs.writeFileSync(row.pkg,result.pkg);row.outputHash=createHash('sha256').update(result.pkg).digest('hex');row.type=result.type;row.stats=result.stats;row.losses=result.losses;
    const check=spawnSync('python3',['-I','tools/ooxml/package.py','check',row.pkg],{encoding:'utf8'});row.status=check.status===0?'converted':'invalid-package';fs.writeFileSync(row.pkg+'.check.txt',check.stdout+check.stderr);
  }catch(e){row.status=['LegacyError','CFBError'].includes(e.name)?'refused':'crashed';row.code=e.code;row.message=e.message;if(row.status==='crashed')row.stack=e.stack;}
  row.ms=performance.now()-start;rows.push(row);console.log(JSON.stringify(row));
}
fs.writeFileSync(path.join(out,'results.json'),JSON.stringify(rows,null,2)+'\n');if(rows.some(r=>['crashed','invalid-package'].includes(r.status)))process.exitCode=1;
