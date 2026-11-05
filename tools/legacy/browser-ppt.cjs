// Browser-only importer and Lectern round trip; Chromium 9223, own contexts only.
const fs=require('node:fs'),path=require('node:path'),{chromium}=require('playwright');
(async()=>{
 const out=path.resolve(process.argv[2]),files=JSON.parse(fs.readFileSync(process.argv[3],'utf8'));fs.mkdirSync(out,{recursive:true});
 const browser=await chromium.connectOverCDP('http://127.0.0.1:9223'),before=await(await fetch('http://127.0.0.1:9223/json')).json(),results=[];
 for(const app of ['quire','ledger','lectern']){
  const context=await browser.newContext({viewport:{width:1280,height:900}}),page=await context.newPage(),errors=[],external=[],samples=[];
  page.on('pageerror',e=>errors.push(e.message));page.on('request',r=>{if(/^https?:/.test(r.url())&&!r.url().startsWith('http://127.0.0.1:8769/'))external.push(r.url());});
  try{
   await page.goto('http://127.0.0.1:8769/'+app+'/');await page.waitForTimeout(600);
   for(const name of ['base','codepages','properties','package','property-parts','images','officeart','native-geometry','shape-path','shapes','xml-check','ppt-read','ppt-structure','ppt-theme','ppt-text-read','ppt-text','ppt-text-design','ppt-text-write','ppt-write'])await page.addScriptTag({url:'/common/legacy/'+name+'.js'});
   for(const [i,file]of(app==='lectern'?files:files.slice(0,1)).entries()){
    const result=await page.evaluate(async({source,extension})=>{const r=await L.legacy.ppt.convert(Uint8Array.from(atob(source),c=>c.charCodeAt(0)),{extension});window.__legacyConversion=r;return {type:r.type,stats:r.stats,bytes:r.pkg.length};},{source:fs.readFileSync('/work/corpora/'+file).toString('base64'),extension:path.extname(file)});
    if(app==='lectern'){
     await page.evaluate(async()=>{const p=await L.pptx.read(window.__legacyConversion.pkg.buffer);L.app.loadPres(p,'Legacy structure');});await page.waitForTimeout(400);
     const saved=await page.evaluate(async()=>{
      const clean=v=>Array.isArray(v)?v.map(clean):v&&typeof v==='object'?Object.fromEntries(Object.keys(v).sort().filter(k=>k!=='keep'&&!k.startsWith('_')).map(k=>[k,clean(v[k])])):v;
      const shapes=list=>(list||[]).map(s=>({type:s.type,ph:clean(s.ph),x:s.x,y:s.y,w:s.w,h:s.h,rot:s.rot,tx:clean(s.tx),children:shapes(s.children)}));
      const snapshot=p=>({width:p.W,height:p.H,firstNum:p.firstNum,slides:p.slides.map(s=>({hidden:s.hidden,notes:s.notes,bg:s.bg,layout:s.layout,design:Object.keys(p.designs).indexOf(s.design),text:L.model.allText(s),shapes:shapes(s.shapes)})),designs:Object.values(p.designs).map(d=>({colors:d.colors,fonts:d.fonts,bg:d.bg,tx:clean(d.tx)}))});
      const before=snapshot(L.pres),bytes=new Uint8Array(await(await L.pptx.write(L.pres,{format:'pptx'})).arrayBuffer()),after=snapshot(await L.pptx.read(bytes.buffer));
      return {before,after,same:JSON.stringify(before)===JSON.stringify(after),bytes:Array.from(bytes)};
     });
     fs.writeFileSync(path.join(out,'lectern-saved-'+i+'.pptx'),Buffer.from(saved.bytes));delete saved.bytes;fs.writeFileSync(path.join(out,'lectern-saved-'+i+'.json'),JSON.stringify(saved,null,2));result.roundtrip=saved.same;result.slides=saved.before.slides.length;
     await page.screenshot({path:path.join(out,'lectern-'+i+'.png')});
    }
    samples.push({file,...result});
   }
   await page.screenshot({path:path.join(out,app+'.png')});results.push({app,samples,errors,external});
  }finally{await context.close();}
 }
 const after=await(await fetch('http://127.0.0.1:9223/json')).json(),retained=before.every(a=>after.some(b=>a.id===b.id));
 const result={results,originalTabs:before.length,retained};fs.writeFileSync(path.join(out,'results.json'),JSON.stringify(result,null,2));console.log(JSON.stringify(result));
 process.exit(results.some(r=>r.errors.length||r.external.length||r.samples.some(s=>s.roundtrip===false))||!retained?1:0);
})().catch(e=>{console.error(e);process.exit(1);});
