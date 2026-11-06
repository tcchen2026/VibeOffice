// Chromium 9223 only; browser Word conversion and Quire save/reopen evidence.
const fs=require('node:fs'),path=require('node:path'),{chromium}=require('playwright');
(async()=>{
 const out=path.resolve(process.argv[2]),files=JSON.parse(fs.readFileSync(process.argv[3],'utf8'));fs.mkdirSync(out,{recursive:true});
 const browser=await chromium.connectOverCDP('http://127.0.0.1:9223'),before=await(await fetch('http://127.0.0.1:9223/json')).json(),results=[];
 for(const app of ['quire','ledger','lectern']){
  const context=await browser.newContext({viewport:{width:1280,height:900}}),page=await context.newPage(),errors=[],external=[],samples=[];
  await context.route('**/common/i18n/**',route=>route.abort());
  page.on('pageerror',e=>errors.push(e.message));page.on('console',m=>{if(m.type()==='error')errors.push(m.text());});page.on('request',r=>{if(/^https?:/.test(r.url())&&!r.url().startsWith('http://127.0.0.1:8769/'))external.push(r.url());});
  try{
   await page.goto('http://127.0.0.1:8769/'+app+'/?new&lang=en');await page.waitForTimeout(600);
   await page.evaluate(()=>L.legacyImport.load());
   for(const [i,file] of (app==='quire'?files:files.slice(0,1)).entries()){
    const result=await page.evaluate(async({source,extension})=>{const r=await L.legacy.doc.convert(Uint8Array.from(atob(source),c=>c.charCodeAt(0)),{extension});window.__legacyConversion=r;return {type:r.type,stats:r.stats,bytes:r.pkg.length};},{source:fs.readFileSync('/work/corpora/'+file).toString('base64'),extension:path.extname(file)});
    if(app==='quire'){
     fs.writeFileSync(path.join(out,'converted-'+i+'.docx'),Buffer.from(await page.evaluate(()=>Array.from(window.__legacyConversion.pkg))));
     await page.evaluate(async()=>{VO.opened=()=>{};const {doc}=await L.docx.read(window.__legacyConversion.pkg.buffer);L.app.loadDoc(doc,'Legacy Word',{saved:true});await document.fonts.ready;L.layout.render();});
     await page.waitForTimeout(900);
     const saved=await page.evaluate(async()=>{
      // Property x.fragment is raw-XML preservation bookkeeping with a fresh
      // package GUID on reopen, not the modeled text/formatting being compared.
      const clean=v=>Array.isArray(v)?v.map(clean):v&&typeof v==='object'?Object.fromEntries(Object.keys(v).sort().filter(k=>!['keep','id','part','dirty'].includes(k)&&!k.startsWith('_')&&!(k==='x'&&v[k]?.fragment)).map(k=>[k,clean(v[k])])):v;
      const snapshot=d=>({stories:[...L.D.stories(d)].map(s=>{const paragraphs=[];L.D.walk(s,b=>{if(b.t==='p')paragraphs.push({text:L.D.plainText(b),pPr:clean(b.pPr),rPr:clean(b.rPr),runs:b.runs.map(r=>clean(r))});return true;});return {kind:s.kind,paragraphs};}),compat:d.settings.compat});
      const before=snapshot(L.D.doc),pages=L.layout.pages.length,bytes=new Uint8Array(await(await L.docx.write(L.D.doc,{format:'docx'})).arrayBuffer()),read=await L.docx.read(bytes.buffer),after=snapshot(read.doc);
      return {before,after,pages,same:JSON.stringify(before)===JSON.stringify(after),bytes:Array.from(bytes),warnings:read.warnings};
     });
     fs.writeFileSync(path.join(out,'quire-saved-'+i+'.docx'),Buffer.from(saved.bytes));delete saved.bytes;fs.writeFileSync(path.join(out,'quire-saved-'+i+'.json'),JSON.stringify(saved,null,2));result.roundtrip=saved.same;result.pages=saved.pages;result.warnings=saved.warnings;
     await page.screenshot({path:path.join(out,'quire-'+i+'.png')});
    }
    samples.push({file,...result});
   }
   await page.screenshot({path:path.join(out,app+'.png')});results.push({app,samples,errors,external});
  }finally{await context.close();}
 }
 const after=await(await fetch('http://127.0.0.1:9223/json')).json(),retained=before.every(a=>after.some(b=>a.id===b.id)),result={results,originalTabs:before.length,retained};
 fs.writeFileSync(path.join(out,'results.json'),JSON.stringify(result,null,2));console.log(JSON.stringify(result));process.exit(results.some(r=>r.errors.length||r.external.length||r.samples.some(s=>s.roundtrip===false))||!retained?1:0);
})().catch(e=>{console.error(e);process.exit(1);});
