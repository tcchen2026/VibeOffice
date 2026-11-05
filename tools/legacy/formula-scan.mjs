// Office-found formula constraints, across measured packages and existing OOXML.
// Diagnostic only: package.py rules belong to Lead. Never execute formulas.
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import {load} from './load.mjs';
const B=load();for(const f of ['public/common/xml.js','public/ledger/js/formula.js'])vm.runInThisContext(fs.readFileSync(f,'utf8'),{filename:f});
const [run,outArg]=process.argv.slice(2),out=outArg||'tmp/legacy/artifacts/C2-formula-scan.json';
const rows=fs.readFileSync(path.join(run,'runs.jsonl'),'utf8').trim().split('\n').map(JSON.parse).filter(r=>r.status==='converted'),files=[];
for(const row of rows)files.push({group:'legacy',file:row.pkg||path.join(run,'packages',row.sha256+(row.mutation?'-'+row.mutation:'')+'.'+row.type),source:row.file,mutation:row.mutation});
for(const file of fs.readdirSync('/work/corpora/excel'))if(/\.(?:xlsx|xlsm|xltx|xltm)$/i.test(file))files.push({group:'ooxml',file:path.join('/work/corpora/excel',file)});
const min=new Map(Object.values(B.xls.functions).map(f=>[f[0],f[1]])),result={files:{},unreadable:[],issues:[],unparsedFormulas:0},selected=new Set();
for(const file of files){
  result.files[file.group]=(result.files[file.group]||0)+1;const issues=[];
  try{
    const zip=await L.zip.read(new Uint8Array(fs.readFileSync(file.file)));
    if(!zip.has('xl/workbook.xml'))throw Error('No SpreadsheetML workbook part');
    for(const [part,entry]of zip){
      if(!/^xl\/(?:worksheets\/[^/]+\.xml|workbook\.xml|externalLinks\/[^/]+\.xml)$/.test(part))continue;
      if(entry.uncompressedSize>128*1024*1024)throw Error('Part exceeds scan limit');
      const xml=await entry.text();
      if(part.startsWith('xl/externalLinks/')){if(/<(?:\w+:)?sheetDataSet(?:\s[^>]*)?\s*\/>|<(?:\w+:)?sheetDataSet(?:\s[^>]*)?>\s*<\/(?:\w+:)?sheetDataSet>/.test(xml))issues.push({part,rule:'external-empty-sheet-data-set'});continue;}
      const re=/<(?:\w+:)?(f|definedName)(\s[^>]*|)>([^<]*)<\/(?:\w+:)?\1>/g;let match;
      while((match=re.exec(xml))){const text=L.xmlTree.unescape(match[3]);if(!text.trim())continue;let ast;
        try{ast=L.formula.parse(text);}catch{result.unparsedFormulas++;continue;}
        const stack=[ast],seen=new Set();while(stack.length){const n=stack.pop();if(!n||typeof n!=='object'||seen.has(n))continue;seen.add(n);
          if(n.t==='fn'){
            const name=n.name.replace(/^_XLFN\.|^_XLL\./,''),minimum=min.get(name);
            if(minimum!=null&&n.args.length<minimum)issues.push({part,rule:'function-required-arguments',name,arguments:n.args.length,minimum,formula:text});
            if(match[1]==='definedName'&&name==='GET.CELL')issues.push({part,rule:'plain-workbook-XLM-name',name,formula:text});
          }
          for(const v of Object.values(n))if(v&&typeof v==='object')stack.push(...(Array.isArray(v)?v:[v]));
        }
      }
    }
  }catch(error){result.unreadable.push({...file,error:error.message});}
  if(issues.length){result.issues.push({...file,issues});if(file.source)selected.add(file.source);}
}
result.counts={};for(const row of result.issues)for(const issue of row.issues){const key=row.group+':'+issue.rule;result.counts[key]=(result.counts[key]||0)+1;}
fs.writeFileSync(out,JSON.stringify(result,null,2)+'\n');fs.writeFileSync(out+'.files.json',JSON.stringify([...selected].sort(),null,2)+'\n');
console.log(JSON.stringify({files:result.files,unreadable:result.unreadable.length,affectedFiles:result.issues.length,counts:result.counts,unparsedFormulas:result.unparsedFormulas,legacySources:selected.size},null,2));
