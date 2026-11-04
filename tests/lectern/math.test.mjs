// Real Office-authored equations are opaque text atoms, including after SmartArt conversion.
import { after, before, test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { ensureServer, openPage } from '../../tools/ooxml/cdp.mjs';
let page, stop;
before(async()=>{stop=await ensureServer();page=await openPage('lectern',{persistence:false,automated:true});});
after(async()=>{try{await page?.close();}finally{stop?.();}});
const input=process.env.MATH_INPUT || path.join(process.env.VO_CORPORA || '/work/corpora','libreoffice/sd/qa/unit/data/pptx/tdf131553.pptx');
const evaluate=async(fn,arg)=>page.evaluate(`(${fn.toString()})(${JSON.stringify(arg)})`);

test('diagram equations survive conversion, adjacent text, copies, history and drafts as opaque atoms',async t=>{
  if(!fs.existsSync(input)) return t.skip('Fetch pinned corpus: '+input);
  const r=await evaluate(async data=>{
    const bytes=Uint8Array.from(atob(data),c=>c.charCodeAt(0));
    L.app.loadPres(await L.pptx.read(bytes),'Equation',{saved:true});L.hist.clear();
    const find=()=>{let match;for(const slide of L.pres.slides)L.model.walk(slide.shapes,sh=>{if(sh.tx?.ps.some(p=>p.rs.some(r=>r.fld==='math')))match||={slide,sh};return true;});return match;};
    let found=find();if(!found)return {missing:true};
    const sourceMath=found.sh.tx.ps.flatMap(p=>p.rs).find(r=>r.fld==='math').keep.math;
    const value=e=>JSON.stringify([e.namespaceURI,e.localName,Array.from(e.attributes).filter(a=>a.name!=='xmlns'&&a.prefix!=='xmlns').map(a=>[a.namespaceURI,a.localName,a.value]).sort(),e.children.length?e.children.map(value):e.textContent]);
    const wanted=value(L.opc.parse(sourceMath.xml));
    const save=async()=>{
      const blob=await L.pptx.write(L.pres),bytes=new Uint8Array(await blob.arrayBuffer()),pkg=await L.opc.open(bytes),math=[];
      for(const part of pkg.names.filter(n=>/^ppt\/slides\/slide\d+\.xml$/.test(n))) for(const m of pkg.xml(part).getElementsByTagNameNS('http://schemas.microsoft.com/office/drawing/2010/main','m')) math.push(value(m));
      let str='';for(let i=0;i<bytes.length;i+=32768)str+=String.fromCharCode(...bytes.subarray(i,i+32768));
      return {bytes:btoa(str),math,parts:pkg.names,losses:blob.dropped};
    };
    const original=await save();
    L.hist.push('Move diagram component');found.sh.x+=6;
    const moved=await save();L.hist.doUndo();const undo=await save();L.hist.doRedo();const redo=await save();
    const top=L.pres.slides.find(s=>s.id===find().slide.id).shapes.find(s=>{let has=false;L.model.walk([s],sh=>{if(sh===find().sh)has=true;return true;});return has;});
    const packet=JSON.parse(JSON.stringify(L.preserve.clipboard({kind:'shapes',shapes:[L.model.dup(top)],text:''})));
    const draftBlob=await __corpusHooks.snapshot(L.pres),info=__corpusHooks.docInfo();L.hist.clear();
    await __corpusHooks.open(new File([draftBlob],'Equation.pptx'),{draft:true,name:'Equation.pptx',saved:true,lossState:info.lossState});
    const draft=await save();
    const recovered=L.pres;
    const copied=L.model.newPresentation();copied.slides[0].shapes=L.preserve.pasteboard(packet).shapes;L.pres=copied;
    const copy=await save();L.pres=recovered;L.app.loadPres(recovered,'Equation',{saved:true});L.hist.clear();
    found=find();L.ed.goto(L.pres.slides.indexOf(found.slide),{force:true});L.ed.render();
    const pi=found.sh.tx.ps.findIndex(p=>p.rs.some(r=>r.fld==='math'));
    L.te.begin(found.sh.id,{sel:{p0:pi,o0:0,p1:pi,o1:0}});
    window.__equationTest={wanted,save,find,pi};
    return {wanted,original,moved,undo,redo,draft,copy,atomic:L.te.state?.root.querySelector('[data-fld="math"]')?.contentEditable};
  },fs.readFileSync(input).toString('base64'));
  assert.equal(r.missing,undefined,'reader creates an equation run');assert.equal(r.atomic,'false');
  for(const k of ['moved','redo','draft','copy'])assert.ok(r[k].math.includes(r.wanted),k+' retains the whole equation');
  assert.deepEqual(r.undo.parts,r.original.parts,'undo restores the diagram graph');
  const removed=r.original.parts.filter(n=>!r.moved.parts.includes(n)&&n.startsWith('ppt/diagrams/'));
  assert.ok(removed.length,'converted diagram has parts to account for');
  const notices=r.moved.losses.filter(n=>n.notify&&n.action==='conversion'&&n.what.includes('IntelliArt'));
  const accounted=new Set(notices.flatMap(n=>JSON.parse(n.detail).removedParts));
  for(const part of removed)assert.ok(accounted.has(part),'shown conversion notice accounts for '+part);
  const pause=()=>new Promise(resolve=>setTimeout(resolve,180));
  await page.send('Input.insertText',{text:'before '});await pause();
  await page.evaluate('L.te.sync();{const q=__equationTest,sh=q.find().sh,n=L.txt.len(sh.tx.ps[q.pi]);L.te.setSel({p0:q.pi,o0:n,p1:q.pi,o1:n});}');
  await page.send('Input.insertText',{text:' after'});await pause();
  const edited=await page.evaluate(`(async()=>{L.te.end();const q=__equationTest,sh=q.find().sh,ps=sh.tx.ps;return {saved:await q.save(),text:L.txt.plain(sh.tx),atoms:ps.flatMap(p=>p.rs).filter(r=>r.fld==='math').length,leaked:ps.flatMap(p=>p.rs).some(r=>r.fld!=='math'&&r.keep?.math)};})()`);
  assert.ok(edited.saved.math.includes(r.wanted));assert.equal(edited.atoms,1);assert.equal(edited.leaked,false);
  assert.match(edited.text,/before .* after/s);
  assert.deepEqual(page.errors,[]);
  if(process.env.MATH_ARTIFACTS){fs.mkdirSync(process.env.MATH_ARTIFACTS,{recursive:true});for(const [k,v]of Object.entries({moved:r.moved,copy:r.copy,draft:r.draft,adjacent:edited.saved}))fs.writeFileSync(path.join(process.env.MATH_ARTIFACTS,k+'.pptx'),Buffer.from(v.bytes,'base64'));}
});

test('equations survive data-only, cache-only and uncached diagram text',async t=>{
  if(!fs.existsSync(input))return t.skip('Fetch pinned corpus: '+input);
  const r=await evaluate(async data=>{
    const bytes=Uint8Array.from(atob(data),c=>c.charCodeAt(0)),z=await L.zip.read(bytes),out=[];
    for(const mode of ['data-only','cache-only','uncached']){
      const entries=[];
      for(const item of new Set(z.values())){
        let data=await item.bytes();
        if(/^ppt\/diagrams\/.*\.xml$/.test(item.name)){
          let xml=await item.text();
          if(mode==='data-only'&&/\/drawing/.test(item.name)||mode==='cache-only'&&/\/data/.test(item.name))xml=xml.replace(/<a14:m\b[^>]*>[\s\S]*?<\/a14:m>/g,'');
          if(mode==='uncached'&&/\/drawing/.test(item.name))xml=xml.replace(/<dsp:spTree\b[^>]*>[\s\S]*?<\/dsp:spTree>/g,'<dsp:spTree/>');
          data=xml;
        }
        entries.push({name:item.name,data});
      }
      L.pres=await L.pptx.read(new Uint8Array(await (await L.zip.write(entries)).arrayBuffer()));
      let atom=0;for(const sl of L.pres.slides)L.model.walk(sl.shapes,sh=>{if(sh.tx?.ps.some(p=>p.rs.some(r=>r.fld==='math'))){atom++;sh.x+=6;}return true;});
      const pkg=await L.opc.open(new Uint8Array(await (await L.pptx.write(L.pres)).arrayBuffer()));
      const slideMath=pkg.names.filter(n=>/^ppt\/slides\/slide\d+\.xml$/.test(n)).reduce((n,p)=>n+pkg.xml(p).getElementsByTagNameNS('http://schemas.microsoft.com/office/drawing/2010/main','m').length,0);
      out.push({mode,atom,slideMath});
    }
    return out;
  },fs.readFileSync(input).toString('base64'));
  for(const v of r){assert.equal(v.atom,1,v.mode+' import');assert.equal(v.slideMath,1,v.mode+' converted output');}
  assert.deepEqual(page.errors,[]);
});
