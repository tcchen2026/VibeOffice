// Property edits, detached clipboard transfer and suite draft recovery.
// node tools/lectern/test/properties.mjs OUTPUT_DIR [SCENARIO_REGEX]
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { ensureServer, openPage } from '../../ooxml/cdp.mjs';
const [out, filter = ''] = process.argv.slice(2);
if (!out) throw Error('Usage: properties.mjs OUTPUT_DIR [SCENARIO_REGEX]');
const corpus = path.join(process.env.VO_CORPORA || path.join(os.homedir(), 'corpora'), 'powerpoint');
const samples = [
 ['shape', 'libreoffice__1643483b371e__effectOrder.pptx', ['shadow', 'copy', 'delete', 'draft']],
 ['text', 'libreoffice__ccdcc5b34e6e__shape-text-glow-effect.pptx', ['bold', 'copy']],
 ['picture', 'libreoffice__3b907cd73bb6__tdf113163.pptx', ['crop', 'paste']],
 ['three', 'libreoffice__021487ac4b55__Scene3d_legacyPerspectiveTopRight.pptx', ['move']],
 ['tags', 'libreoffice__22340d3cfa1d__tdf126324.pptx', ['copy', 'delete', 'paste']],
 ['action', 'libreoffice__8312b8523db8__macro.pptm', ['copy', 'replace']],
 ['transition', 'libreoffice__06c8ac85b26d__bnc904423.pptx', ['advance', 'replace']],
 ['animation', 'openxml-sdk__0e4818677b1c__Text_withEffects_100chars+Animation (Fly In, by letter).pptx', ['animation', 'delete']],
];
const cases = samples.flatMap(([kind,file,edits]) => edits.map(edit => ({file,kind,edit,scenario:kind+'-'+edit})));
const selected = cases.filter(c => new RegExp(filter).test(c.scenario));
fs.mkdirSync(out, { recursive: true });
const results = path.join(out, 'results.jsonl');
const rows = fs.existsSync(results) ? fs.readFileSync(results, 'utf8').trim().split('\n').filter(Boolean).map(JSON.parse).filter(r => !selected.some(c => c.scenario === r.scenario)) : [];
const stop = await ensureServer(), page = await openPage('lectern', { persistence: false });
try {
  for (const c of selected) {
    const row = { ...c, attempted: true }; page.errors.length = 0;
    try {
      const input = path.join(c.multi ? process.env.FRAME_FIXTURES || path.join(path.dirname(out), 'frames-fixtures') : corpus, c.file);
      const result = await page.evaluate(`(${run.toString()})(${JSON.stringify({ ...c, data: fs.readFileSync(input).toString('base64') })})`);
      for (const [state, data] of Object.entries(result.artifacts)) {
        const dir = path.join(out, state === 'saved' ? c.scenario : c.scenario + '-' + state);
        fs.mkdirSync(dir, { recursive: true }); fs.writeFileSync(path.join(dir, c.file), Buffer.from(data, 'base64'));
      }
      delete result.artifacts; Object.assign(row, result);
      if (page.errors.length) throw Error(page.errors.join('\n'));
      if (c.edit === 'draft' || c.edit === 'paste') {
        const shot = await page.send('Page.captureScreenshot', { format: 'png' });
        fs.writeFileSync(path.join(out, c.scenario + '.png'), Buffer.from(shot.data, 'base64'));
      }
    } catch (error) { Object.assign(row, { status: 'failed', error: error.stack }); }
    rows.push(row); fs.writeFileSync(results, rows.map(r => JSON.stringify(r)).join('\n') + '\n');
    console.log(c.scenario + ': ' + row.status + (row.error ? ' ' + row.error : ''));
  }
} finally { await page.close(); stop(); }
if (rows.some(r => r.status !== 'ok')) process.exitCode = 1;

async function run(o) {
  const K=L.opc,M=L.model,A=L.app,H=L.hist,E=L.ed;
  const check=(ok,why)=>{if(!ok)throw Error(why)};
  const all=e=>e?[e,...e.getElementsByTagName('*')]:[];
  const feature={shape:'glow',text:'glow',picture:'clrChange',three:'scene3d',tags:'custDataLst',action:'hlinkClick',transition:'transition',animation:'reflection'}[o.kind];
  const stats=pkg=>{
    let count=0;for(const r of pkg.rels(pkg.main).filter(r=>/\/slide$/.test(r.type)))count+=all(pkg.xml(r.part)).filter(e=>e.localName===feature).length;
    return count;
  };
  VO.opened=()=>{}; H.clear();
  await __corpusHooks.open(new File([Uint8Array.from(atob(o.data),c=>c.charCodeAt(0))],o.file));
  const pres=L.pres,before=stats(pres.pkg),artifacts={};let slide,shape;
  for(const s of pres.slides){
    M.walk(s.shapes,sh=>{
      const pr=sh.keep?.properties, xml=o.kind==='action'?sh.keep?.identity?.xml:JSON.stringify(o.kind==='text'?sh.tx?.ps:o.kind==='tags'?pr?.nonVisual:pr);
      if(!shape && (o.kind==='animation'?s.anims.some(a=>a.sid===sh.id):xml?.includes(feature))){shape=sh;slide=s;}
      return true;
    });if(shape)break;
  }
  if(o.kind==='transition'){slide=pres.slides.find(s=>s.keep?.transition);shape=slide.shapes[0];}
  check(slide,'No target '+o.kind); E.goto(pres.slides.indexOf(slide),{force:true});A.view='normal';A.focusArea='editor';E.select(shape?[shape.id]:[]);H.clear();
  let expected=before;
  const piece=shape?all(K.parse('<root xmlns:p="'+K.NS.p+'" xmlns:a="'+K.NS.a+'">'+(shape.keep?.properties?.spPr?.xml||'')+(shape.keep?.properties?.blipFill?.xml||'')+(shape.keep?.identity?.xml||'')+(shape.keep?.properties?.nonVisual||[]).map(f=>f.xml).join('')+'</root>')).filter(e=>e.localName===feature).length:0;
  const ownRunCount=shape?.tx?.ps.reduce((n,p)=>n+p.rs.reduce((a,r)=>a+(r.keep?.fragment?.xml.match(new RegExp('<[^/<>:]+:'+feature+'(?=[ >])','g'))||[]).length,0),0)||0;
  async function save(state,blob){
    blob ||= await L.pptx.write(L.pres); const pkg=await K.open(new Uint8Array(await blob.arrayBuffer()));
    if(o.kind!=='animation')check(stats(pkg)===expected,'Property count '+JSON.stringify({actual:stats(pkg),expected,losses:L.pres.losses}));
    if(o.kind==='tags'&&['copy','paste'].includes(o.edit)&&state!=='undo'){
      const refs=pkg.rels(slide.keep.part).filter(r=>/\/tags$/.test(r.type));check(new Set(refs.map(r=>r.part)).size>=2,'Copied tags share a part');
    }
    if(o.kind==='transition'&&o.edit==='advance'&&state!=='undo'){
      const transitions=all(pkg.xml(slide.keep.part)).filter(e=>e.localName==='transition');check(transitions.every(e=>e.getAttribute('advTm')==='12345'),'Advance edit missing');
      check(transitions.some(e=>Array.from(e.attributes).some(a=>a.localName==='dur')),'Extended transition duration lost');
    }
    const bytes=new Uint8Array(await blob.arrayBuffer());let text='';for(let i=0;i<bytes.length;i+=32768)text+=String.fromCharCode(...bytes.subarray(i,i+32768));artifacts[state]=btoa(text);
  }
  if(o.edit==='bold')A.toggleRun('b');
  else if(o.edit==='paste'){
    const pasted=L.preserve.pasteboard(JSON.parse(JSON.stringify(L.preserve.clipboard({kind:'shapes',shapes:[L.clone(shape)],text:''}))));
    const images=[];M.walk(pasted.shapes,s=>{if(s.media)images.push(s.media);return true});
    for(const id of images){check(L.media.has(id),'Clipboard preview missing');await L.loadImage(L.media.url(id));}
    A.pasteItem(pasted);
  }
  else {
    H.push('Property '+o.edit);
    if(o.edit==='copy')E.find(shape.id).list.push(M.dup(shape));
    else if(o.edit==='delete'){const list=E.find(shape.id).list;list.splice(list.indexOf(shape),1);slide.anims=slide.anims.filter(a=>a.sid!==shape.id);}
    else if(o.edit==='shadow'||o.edit==='draft')shape.shadow={c:'#225588',a:.4,dx:5,dy:4,blur:3};
    else if(o.edit==='crop')shape.crop={...shape.crop,l:.1};
    else if(o.edit==='move')M.translate(shape,12,8);
    else if(o.edit==='advance')slide.trans.after=12345;
    else if(o.kind==='transition')slide.trans={type:'fade',spd:'fast',click:true,after:null};
    else if(o.kind==='action')shape.link={url:'https://example.invalid/replaced'};
    else if(o.edit==='animation')slide.anims.find(a=>a.sid===shape.id).dur+=200;
  }
  if(['copy','paste'].includes(o.edit))expected+=piece+ownRunCount;
  if(o.edit==='delete')expected-=piece+ownRunCount;
  if(o.kind==='transition'&&o.edit==='replace')expected=before-all(K.parse(slide.keep.transition.xml)).filter(e=>e.localName==='transition').length+1;
  const after=expected;
  await save('saved');check(H.doUndo(),'Undo missing');expected=before;await save('undo');check(H.doRedo(),'Redo missing');expected=after;await save('redo');
  if(o.edit==='draft'){
    const info=__corpusHooks.docInfo(L.pres),blob=await __corpusHooks.snapshot(L.pres);await save('draft',blob);H.clear();
    await __corpusHooks.open(new File([blob],info.draftName||info.name),{draft:true,name:info.name,saved:true,lossState:info.lossState});await save('recovered');
  }
  return {status:'ok',artifacts,before,after,expectedSlides:L.pres.slides.length,losses:L.pres.losses};
}
