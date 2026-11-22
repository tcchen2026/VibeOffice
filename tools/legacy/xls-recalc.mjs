// Recalculate converted workbooks with the real Ledger engine. Every BIFF formula
// cell, including array/table members, is counted against its recorded cache.
import fs from 'node:fs';
import path from 'node:path';
import {createRequire} from 'node:module';
import {Worker,isMainThread,parentPort,workerData} from 'node:worker_threads';
const root=path.resolve(import.meta.dirname,'../..');
if(!isMainThread){
  const require=createRequire(import.meta.url),L=require('../../tests/ledger/load.js'),F=L.formula,M=L.model;
  const out={file:workerData.file,source:0,compared:0,agree:0,excluded:{},functions:{},mismatches:[],flags:{}};
  const errors={0:'#NULL!',7:'#DIV/0!',15:'#VALUE!',23:'#REF!',29:'#NAME?',36:'#NUM!',42:'#N/A',43:'#GETTING_DATA'};
  const same=(a,b)=>a===b||(M.isErr(a)&&M.isErr(b)&&a.e===b.e)||(typeof a==='number'&&typeof b==='number'&&(Math.abs(a-b)<=Math.max(Math.abs(a),Math.abs(b))*1e-9||Math.abs(a-b)<1e-10));
  const exclude=reason=>out.excluded[reason]=(out.excluded[reason]||0)+1;
  try{
    const start=performance.now(),wb=await L.xlsxRead.read(fs.readFileSync(workerData.pkg));out.readMs=performance.now()-start;
    wb.fileName=path.basename(workerData.file).replace(/^[a-z]+__[0-9a-f]+__/,'');out.calcPr=wb.calcPr;
    const refs=JSON.parse(fs.readFileSync(workerData.reference,'utf8')).references.sheets,expected=[];
    for(const source of refs){
      const sh=wb.sheets.find(s=>s.name===source.name),groups=[];
      if(sh)sh.rows.forEach(row=>row?.cells.forEach(cell=>{if(cell?.af)groups.push({cell,range:cell.af});else if(cell?.dt){const range=F.parseRange(cell.dt.ref);if(range)groups.push({cell,range});}}));
      for(const c of source.cells){if(!('formula' in c))continue;out.source++;
        const cell=sh?.get(c.row,c.col),group=groups.find(g=>c.row>=g.range.r1&&c.row<=g.range.r2&&c.col>=g.range.c1&&c.col<=g.range.c2),owner=group?.cell||cell;
        if(!owner){exclude('missingCell');continue;}
        if(owner.dt){exclude('dataTablePreserved');continue;}
        if(owner.f==null){exclude('formulaDropped');continue;}
        const value=c.type==='e'?M.err(errors[c.value]||String(c.value)):c.type==='b'?!!c.value:c.type==='empty'?'':c.value;
        expected.push({sh,c,owner,f:owner.f,value});
      }
    }
    const startCalc=performance.now();L.calc.recalcAll(wb);out.calcMs=performance.now()-startCalc;
    const functionNames=new Map();
    const functionsOf=owner=>{
      if(functionNames.has(owner))return functionNames.get(owner);
      const names=new Set(),seen=new Set(),stack=[L.calc.astOf(owner)];
      while(stack.length){const n=stack.pop();if(!n||typeof n!=='object'||seen.has(n))continue;seen.add(n);if(n.t==='fn')names.add(n.name);for(const value of Object.values(n))if(value&&typeof value==='object')stack.push(...(Array.isArray(value)?value:[value]));}
      const result=[...names];functionNames.set(owner,result);return result;
    };
    for(const e of expected){
      if(e.owner.kept){exclude('enginePreserved');continue;}
      if(L.calc.circular.has(e.owner)){exclude('circular');continue;}
      if(/\b(NOW|TODAY|RAND|RANDBETWEEN|RANDARRAY|INFO)\s*\(/i.test(e.f)||/\bCELL\s*\(\s*"(?:filename|address)"/i.test(e.f)){exclude('volatile');continue;}
      const got=e.sh.get(e.c.row,e.c.col)?.v,ok=same(got,e.value),functions=functionsOf(e.owner);
      out.compared++;if(ok)out.agree++;
      if(wb.calcPr.fullCalcOnLoad)out.flags.fullCalcOnLoad=(out.flags.fullCalcOnLoad||0)+1;
      for(const fn of functions){const counts=out.functions[fn]||(out.functions[fn]={agree:0,differ:0});counts[ok?'agree':'differ']++;}
      if(!ok)out.mismatches.push({sheet:e.sh.name,cell:F.cellName(e.c.row,e.c.col),formula:e.f,expected:e.value,actual:got,member:e.owner!==e.sh.get(e.c.row,e.c.col)||undefined,functions});
    }
    parentPort.postMessage(out);
  }catch(error){parentPort.postMessage({...out,error:error.stack});}
}else{
  const run=path.resolve(process.argv[2]),outDir=path.resolve(process.argv[3]||path.join(run,'recalc'));fs.mkdirSync(outDir,{recursive:true});
  const listing=path.join(run,'results.json'),rows=fs.existsSync(listing)?JSON.parse(fs.readFileSync(listing,'utf8')):fs.readFileSync(path.join(run,'runs.jsonl'),'utf8').trim().split('\n').map(JSON.parse);
  const selected=rows.filter(r=>!r.mutation&&r.status==='converted'&&r.stats?.formulas?.source),results=[];let next=0;
  const runOne=row=>new Promise(resolve=>{
    const sha=row.sha256||row.sourceHash,worker=new Worker(new URL(import.meta.url),{workerData:{file:row.file,pkg:row.pkg||path.join(run,'packages',sha+'.'+row.type),reference:path.join(root,'tmp/legacy/artifacts/B-complete/references',sha+'.json')},resourceLimits:{maxOldGenerationSizeMb:1024,stackSizeMb:8}});
    let done=false;const finish=r=>{if(done)return;done=true;clearTimeout(timer);worker.terminate();fs.writeFileSync(path.join(outDir,sha+'.json'),JSON.stringify(r,null,2)+'\n');resolve(r);};
    const timer=setTimeout(()=>finish({file:row.file,error:'TIMEOUT',expectedSource:row.stats.formulas.source}),60000);
    worker.on('message',finish);worker.on('error',e=>finish({file:row.file,error:e.message,expectedSource:row.stats.formulas.source}));worker.on('exit',code=>{if(!done)finish({file:row.file,error:'WORKER_EXIT '+code});});
  });
  if(process.argv.includes('--summarize-only'))for(const row of selected)results.push(JSON.parse(fs.readFileSync(path.join(outDir,(row.sha256||row.sourceHash)+'.json'),'utf8')));
  else await Promise.all(Array.from({length:2},async()=>{while(next<selected.length){const row=await runOne(selected[next++]);results.push(row);if(results.length%25===0)console.log(results.length+'/'+selected.length);}}));
  const summary={files:results.length,failed:results.filter(r=>r.error),source:0,compared:0,agree:0,excluded:{},functions:{},flags:{}};
  for(const r of results){if(r.error)continue;for(const k of ['source','compared','agree'])summary[k]+=r[k];for(const key of ['excluded','flags'])for(const [k,v]of Object.entries(r[key]))summary[key][k]=(summary[key][k]||0)+v;for(const [fn,counts]of Object.entries(r.functions)){const sum=summary.functions[fn]||(summary.functions[fn]={agree:0,differ:0});sum.agree+=counts.agree;sum.differ+=counts.differ;}}
  fs.writeFileSync(path.join(outDir,'summary.json'),JSON.stringify(summary,null,2)+'\n');console.log(JSON.stringify(summary,null,2));if(summary.failed.length)process.exitCode=1;
}
