// Scope of the Office diagnostic's unqualified global-name failure. No evaluation.
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import {load} from './load.mjs';
load();for(const f of ['public/common/xml.js','public/ledger/js/formula.js'])vm.runInThisContext(fs.readFileSync(f,'utf8'));
const run=process.argv[2],out=process.argv[3],rows=fs.readFileSync(path.join(run,'runs.jsonl'),'utf8').trim().split('\n').map(JSON.parse).filter(r=>r.status==='converted');
const files=rows.map(r=>({group:'legacy',source:r.file,mutation:r.mutation,file:r.pkg}));
for(const name of fs.readdirSync('/work/corpora/excel'))if(/\.(?:xlsx|xlsm|xltx|xltm)$/i.test(name))files.push({group:'ooxml',file:path.join('/work/corpora/excel',name)});
const report={files:{},names:0,unparsed:[],unreadable:[],findings:[]};
for(const file of files){
  report.files[file.group]=(report.files[file.group]||0)+1;
  try{
    const zip=await L.zip.read(fs.readFileSync(file.file)),book=zip.get('xl/workbook.xml');if(!book)throw Error('Not SpreadsheetML');
    const root=L.xmlTree.parse(await book.text());
    const pending=[root];while(pending.length){const el=pending.pop();if(typeof el!=='object')continue;if(el.children)pending.push(...el.children);
      if(el.localName!=='definedName'||el.getAttribute('localSheetId')!=null)continue;
      report.names++;const text=el.textContent;let ast;
      try{ast=L.formula.parse('('+text.replace(/^=/,'')+')');}catch(e){report.unparsed.push({...file,name:el.attrs.name,formula:text,error:e.message});continue;}
      const nodes=[ast];let found=false;
      while(nodes.length){const n=nodes.pop();if(!n||typeof n!=='object')continue;
        if(['ref','area'].includes(n.t)&&n.sheet==null&&n.book==null&&!n.bang)found=true;
        for(const v of Object.values(n))if(v&&typeof v==='object')nodes.push(...(Array.isArray(v)?v:[v]));
      }
      if(found)report.findings.push({...file,name:el.attrs.name,formula:text});
    }
  }catch(e){report.unreadable.push({...file,error:e.message});}
}
fs.writeFileSync(out,JSON.stringify(report,null,2)+'\n');console.log(JSON.stringify({files:report.files,names:report.names,unparsed:report.unparsed.length,unreadable:report.unreadable.length,findings:report.findings.length,byGroup:report.findings.reduce((r,v)=>(r[v.group]=(r[v.group]||0)+1,r),{})}));
