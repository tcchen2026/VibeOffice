// Chromium 9223 only. Own contexts/tabs, no app changes or translation inspection.
const fs=require('node:fs'),path=require('node:path'),{chromium}=require('playwright');
(async()=>{
  const out=path.resolve('tmp/legacy/artifacts/browser-C1');fs.mkdirSync(out,{recursive:true});
  const browser=await chromium.connectOverCDP('http://127.0.0.1:9223'),before=await (await fetch('http://127.0.0.1:9223/json')).json(),results=[];
  const source=fs.readFileSync('/work/corpora/legacy-excel/poi__286c52c46a1a__Formatting.xls').toString('base64');
  for(const app of ['quire','ledger','lectern']){
    const context=await browser.newContext({viewport:{width:1280,height:900}}),page=await context.newPage(),errors=[],external=[];
    page.on('pageerror',e=>errors.push(e.message));page.on('request',r=>{if(/^https?:/.test(r.url())&&!r.url().startsWith('http://127.0.0.1:8769/'))external.push(r.url());});
    try{
      await page.goto('http://127.0.0.1:8769/'+app+'/');await page.waitForTimeout(600);
      for(const name of ['base','codepages','properties','package','property-parts','images','officeart','native-geometry','shape-path','shapes','xls-read','xml-check','xls-style-ext','xls-style','xls-sheet','xls-write'])await page.addScriptTag({url:'/common/legacy/'+name+'.js'});
      const result=await page.evaluate(async input=>{
        const B=L.legacy,bytes=Uint8Array.from(atob(input),c=>c.charCodeAt(0)),r=await B.xls.convert(bytes),shape={id:1,type:1,flags:0,properties:new Map([[4,{value:90*65536}],[384,{value:4}]])},rect={x:0,y:0,cx:1000,cy:400};
        if(B.shapes.anchorRect(shape,rect).cx!==400||!B.shapes.style(shape).includes('<a:lin'))throw Error('Shared shape correction failed');
        window.__legacyConversion=r;return {type:r.type,stats:r.stats,bytes:r.pkg.length};
      },source);
      if(app==='ledger'){
        await page.evaluate(async()=>{await L.app.openFile(new File([window.__legacyConversion.pkg],'C1-browser.xlsx',{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));});
        await page.waitForTimeout(400);
        const saved=await page.evaluate(async()=>{
          const snapshot=wb=>wb.sheets.map(s=>({name:s.name,cells:s.rows.flatMap((row,r)=>(row?.cells||[]).flatMap((c,col)=>c?[{r,c:col,v:c.v}]:[])),view:s.view,merges:s.merges}));
          const before=snapshot(L.grid.wb),bytes=await L.xlsxWrite.bytes(L.grid.wb,{type:'xlsx'}),after=snapshot(await L.xlsxRead.read(bytes));
          return {bytes:Array.from(bytes),same:JSON.stringify(before)===JSON.stringify(after),before,after};
        });
        fs.writeFileSync(path.join(out,'ledger-saved.xlsx'),Buffer.from(saved.bytes));delete saved.bytes;fs.writeFileSync(path.join(out,'roundtrip.json'),JSON.stringify(saved,null,2));result.roundtrip=saved.same;
        await page.evaluate(async bytes=>{await L.app.openFile(new File([Uint8Array.from(bytes)],'C1-reopened.xlsx'));},Array.from(fs.readFileSync(path.join(out,'ledger-saved.xlsx'))));
        await page.waitForTimeout(300);
      }
      await page.screenshot({path:path.join(out,app+'.png')});results.push({app,...result,errors,external});
    }finally{await context.close();}
  }
  const after=await (await fetch('http://127.0.0.1:9223/json')).json();const retained=before.every(a=>after.some(b=>a.id===b.id));
  fs.writeFileSync(path.join(out,'results.json'),JSON.stringify({results,originalTabs:before.length,retained},null,2));console.log(JSON.stringify({results,originalTabs:before.length,retained}));
  process.exit(results.some(r=>r.errors.length||r.external.length||r.roundtrip===false)||!retained?1:0);
})().catch(e=>{console.error(e);process.exit(1);});
