// Issue #5: capture Quire's columns/pictures and an unedited save for the Office 2016 API.
// node tests/from-issues/5/inspect.mjs INPUT.docx OUTPUT_DIR
import fs from 'node:fs';
import path from 'node:path';
import { createHash } from 'node:crypto';
import { execFileSync } from 'node:child_process';
import { ensureServer, openPage } from '../../../tools/ooxml/cdp.mjs';

const [source, output] = process.argv.slice(2);
if (!source || !output) throw new Error('Usage: inspect.mjs INPUT.docx OUTPUT_DIR');
fs.mkdirSync(output, { recursive: true });
const bytes = fs.readFileSync(source);
const stop = await ensureServer();
const page = await openPage('quire', { persistence: false, automated: true });
try {
  await page.send('Emulation.setDeviceMetricsOverride', { width: 1440, height: 1500, deviceScaleFactor: 1, mobile: false });
  const result = await page.evaluate(`(async () => {
    VO.opened = () => {};
    L.layout.view = 'print'; L.layout.setZoom(1);
    const {doc} = await L.docx.read(Uint8Array.from(atob(${JSON.stringify(bytes.toString('base64'))}), c => c.charCodeAt(0)));
    L.app.loadDoc(doc, 'Issue 5', {saved: true});
    await document.fonts.ready;
    await Promise.all([...document.images].map(image => image.decode().catch(() => {})));
    L.layout.render();
    await new Promise(resolve => setTimeout(resolve, 350));
    const images = [];
    L.D.walk(doc.main, p => {
      if (p.t === 'p') for (const run of p.runs) if (run.t === 'img') images.push({id:run.id, paragraph:p.id, width:run.w, height:run.h, float:run.float || null});
    });
    const pages = L.layout.pages.map(pg => {
      const origin = pg.el.getBoundingClientRect(), scale = L.layout.scale(), factor = 72/96/scale;
      const rect = el => {const r=el.getBoundingClientRect(); return {x:(r.x-origin.x)*factor,y:(r.y-origin.y)*factor,width:r.width*factor,height:r.height*factor};};
      return {page:pg.idx+1, size:rect(pg.el), text:pg.body.innerText,
        pictures:[...pg.body.querySelectorAll('img')].map(el => ({...rect(el),naturalWidth:el.naturalWidth,naturalHeight:el.naturalHeight,src:el.currentSrc.startsWith('data:')?'data:':el.currentSrc,style:el.getAttribute('style'),parentClass:el.parentElement.className}))};
    });
    const blob = await L.docx.write(doc, {stats:L.app.docStats()});
    const saved = new Uint8Array(await blob.arrayBuffer());
    let binary=''; for(let i=0;i<saved.length;i+=32768) binary+=String.fromCharCode(...saved.subarray(i,i+32768));
    return {pages,images,section:doc.sect,losses:doc.losses || [],saved:btoa(binary)};
  })()`);
  const saved = Buffer.from(result.saved, 'base64'); delete result.saved;
  fs.writeFileSync(path.join(output, 'quire-saved.docx'), saved);
  for (let i = 0; i < result.pages.length; i++) {
    await page.evaluate(`L.layout.pages[${i}].el.scrollIntoView({block:'start'})`);
    await new Promise(resolve => setTimeout(resolve, 150));
    const clip = await page.evaluate(`(() => {const r=L.layout.pages[${i}].el.getBoundingClientRect(); return {x:r.x+scrollX,y:r.y+scrollY,width:r.width,height:r.height,scale:1};})()`);
    const shot = await page.send('Page.captureScreenshot', {format:'png',captureBeyondViewport:true,clip});
    fs.writeFileSync(path.join(output, `quire-page-${i+1}.png`), Buffer.from(shot.data, 'base64'));
  }
  Object.assign(result, {
    source:path.resolve(source), sourceSha256:createHash('sha256').update(bytes).digest('hex'),
    savedSha256:createHash('sha256').update(saved).digest('hex'),
    commit:execFileSync('git',['rev-parse','HEAD'],{encoding:'utf8'}).trim(),
    errors:page.errors,externalRequests:page.requests,windows:page.windows,navigations:page.navigations
  });
  fs.writeFileSync(path.join(output, 'quire.json'), JSON.stringify(result,null,2)+'\n');
  if (page.errors.length || page.requests.length || page.windows.length || page.navigations.length) throw new Error('Browser check failed; see quire.json');
  console.log(JSON.stringify({pages:result.pages.length,images:result.images,losses:result.losses,output},null,2));
} finally { await page.close(); stop(); }
