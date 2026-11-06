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
   await page.evaluate(()=>L.legacyImport.load());
   for(const [i,file]of(app==='lectern'?files:files.slice(0,1)).entries()){
    const result=await page.evaluate(async({source,extension})=>{const r=await L.legacy.ppt.convert(Uint8Array.from(atob(source),c=>c.charCodeAt(0)),{extension});window.__legacyConversion=r;return {type:r.type,stats:r.stats,bytes:r.pkg.length};},{source:fs.readFileSync('/work/corpora/'+file).toString('base64'),extension:path.extname(file)});
    if(app==='lectern'){
     await page.evaluate(async()=>{const p=await L.pptx.read(window.__legacyConversion.pkg.buffer);L.app.loadPres(p,'Legacy structure');});await page.waitForTimeout(400);
     const saved=await page.evaluate(async()=>{
      let media=new Map(),slideIds=new Map();
      const clean=(v,key)=>Array.isArray(v)?v.map(x=>clean(x)):v&&typeof v==='object'?Object.fromEntries(Object.keys(v).sort().filter(k=>k!=='keep'&&!k.startsWith('_')).map(k=>[k,clean(v[k],k)])):key==='media'&&media.has(v)?media.get(v):key==='slide'&&slideIds.has(v)?'slide '+slideIds.get(v):v;
      const shapes=list=>(list||[]).map(s=>({type:s.type,geom:s.geom,path:clean(s.path),ph:clean(s.ph),x:s.x,y:s.y,w:s.w,h:s.h,rot:s.rot==null?s.rot:+(((s.rot%360)+360)%360).toFixed(9),flipH:s.flipH,flipV:s.flipV,tx:clean(s.tx),fill:clean(s.fill),line:clean(s.line),shadow:clean(s.shadow),crop:clean(s.crop),img:clean(s.img),media:clean(s.media,'media'),tbl:clean(s.tbl),link:clean(s.link),children:shapes(s.children)}));
      const snapshot=async p=>{
       slideIds=new Map(p.slides.map((s,i)=>[s.id,i+1]));
       media=new Map(await Promise.all(Object.keys(p.keep?.media||{}).map(async id=>{const b=await L.media.get(id).blob.arrayBuffer(),hash=Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256',b)),v=>v.toString(16).padStart(2,'0')).join('');return[id,hash];})));
       return {width:p.W,height:p.H,firstNum:p.firstNum,slides:p.slides.map(s=>({hidden:s.hidden,notes:s.notes,bg:clean(s.bg),layout:s.layout,design:Object.keys(p.designs).indexOf(s.design),text:L.model.allText(s),transition:clean(s.trans),shapes:shapes(s.shapes)})),designs:Object.values(p.designs).map(d=>({colors:d.colors,fonts:d.fonts,bg:clean(d.bg),tx:clean(d.tx)}))};
      };
      const before=await snapshot(L.pres),bytes=new Uint8Array(await(await L.pptx.write(L.pres,{format:'pptx'})).arrayBuffer()),after=await snapshot(await L.pptx.read(bytes.buffer));
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
