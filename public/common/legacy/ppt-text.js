/* Text shapes, outline references and master inheritance for binary PowerPoint.
 * Derived from Apache POI hslf/usermodel/{HSLFTextShape,HSLFTextParagraph}.java,
 * hslf/record/PPDrawing.java and sl/usermodel/Placeholder.java (Apache-2.0).
 * Derived from Apache OpenOffice filter/source/msfilter/{svdfppt,msdffimp}.cxx (Apache-2.0).
 * MS-PPT §§2.7, 2.9: source shape IDs/placeholder positions remain identities. */
(function(root){
  'use strict';
  const B=root.L.legacy,P=B.ppt,T=P.text,A=B.officeart,N=s=>s;
  const messages={format:N('Some text formatting could not be kept'),text:N('Some presentation text could not be kept'),reference:N('Text with a missing outline reference could not be kept'),drawing:N('Some text-box positioning could not be kept'),bullet:N('Picture bullets were replaced with text bullets'),field:N('A date field format could not be kept')};
  T.reporter=function(book,page){const seen=new Map();return (kind,detail)=>{if(!seen.has(kind)){const loss=B.loss('legacy-presentation-'+kind,messages[kind]||messages.format,page.where,'drop',page.where,true);loss.detail=[];book.losses.push(loss);seen.set(kind,loss);}const list=seen.get(kind).detail;if(list.length<12)list.push(String(detail?.code||detail||kind));};};
  const defaults=type=>({p:{bullet:[1,5,7,8].includes(type),bulletChar:0x2022,bulletHasFont:false,bulletHasColor:false,bulletHasSize:false,align:0,lineSpacing:100,spaceBefore:[1,5,7,8].includes(type)?20:type===2?30:0,spaceAfter:0,margin:0,indent:0,defaultTab:576,rtl:false},r:{font:0,size:[0,6].includes(type)?44:type===2?12:[1,5,7,8].includes(type)?32:24,color:([0,6].includes(type)?3:1)*0x1000000,bold:false,italic:false,underline:false,baseline:0}});
  const baseType=type=>({5:1,6:0,7:1,8:1})[type];
  T.style=function(book,page,type,level){
    type=type>=0&&type<=8?type:4;level=Math.min(8,Math.max(0,level));
    page||=book.defaultTextPage;const key=type+':'+level;if(page.styleCache.has(key))return page.styleCache.get(key);
    const chain=[],seen=new Set();let source=page;
    while(source&&!seen.has(source)&&chain.length<B.limits.depth){seen.add(source);chain.unshift(source);source=source.kind==='notes'?book.notesMaster:source.parent;}
    const d=defaults(type),out={p:{...d.p},r:{...d.r}},base=baseType(type);
    const apply=(map,t)=>{const levels=map?.get(t);if(!levels)return;let previous={p:{},r:{}};for(let i=0;i<=level;i++){const s=levels[i];if(s)previous={p:{...previous.p,...s.p},r:{...previous.r,...s.r}};}Object.assign(out.p,previous.p);Object.assign(out.r,previous.r);};
    apply(book.defaultStyles,4);if(base!=null)apply(book.defaultStyles,base);if(type!==4)apply(book.defaultStyles,type);
    if(type===2&&book.masters[0])apply(book.masters[0].textStyles,2);
    // Environment's other-text style takes priority over a master other style.
    if(type!==4){for(const p of chain){if(base!=null)apply(p.textStyles,base);apply(p.textStyles,type);}}
    else if(!book.defaultStyles.has(4))for(const p of chain)apply(p.textStyles,4);
    page.styleCache.set(key,out);return out;
  };
  T.placeholder=function(rec,extension){
    if(!rec&&!extension)return null;const r=rec?.r,position=r?r.i32(0):-1,id=extension?extension.r.u8(0):r.u8(4),size=r?r.u8(5):undefined;
    const type={1:'title',2:'body',3:'ctrTitle',4:'subTitle',5:'sldImg',6:'body',7:'dt',8:'sldNum',9:'ftr',10:'hdr',11:'sldImg',12:'body',13:'title',14:'body',15:'ctrTitle',16:'subTitle',17:'title',18:'body',19:'obj',20:'chart',21:'tbl',22:'clipArt',23:'dgm',24:'media',25:'obj',26:'pic'}[id];
    return type?{type,idx:position>=0?position:undefined,sz:['full','half','quarter'][size],orient:[17,18,25].includes(id)?'vert':undefined,nativeId:id}:null;
  };
  T.prepare=function(book){
    const pages=[...book.masters,...book.slides,...book.notes,book.notesMaster,book.handoutMaster].filter(Boolean);
    book.defaultTextPage={styleCache:new Map()};book.defaultStyles=new Map();book.textStats={textShapes:0,textCharacters:0,paragraphs:0,fields:0,placeholders:0};
    let recordBudget=B.limits.records,textBudget=B.limits.text;
    const take=()=>{if(--recordBudget<0)B.fail('PPT_TEXT_LIMIT',N('The presentation has too many text records'));};
    const env=book.documentRecords.find(r=>r.type===1010),envRecords=env?P.children(env):[];
    const envReport=T.reporter(book,{where:'Presentation'});
    const readStyles=(records,map,report)=>{for(const rec of records)if(rec.type===4003){try{map.set(rec.instance,T.master(rec));}catch(e){if(e.name!=='LegacyError')throw e;report('format',e);}}};
    readStyles(envRecords,book.defaultStyles,envReport);
    book.defaultSI={};for(const rec of envRecords)if(rec.type===4009)try{book.defaultSI=T.si(T.cursor(rec.r));}catch(e){if(e.name!=='LegacyError')throw e;}
    const headers=(records,report)=>{const out={};for(const parent of records.filter(r=>r.type===4057)){const kids=P.children(parent,e=>report('field',e));for(const r of kids){if(r.type===4058&&r.r.length>=4){out.format=r.r.u16(0);out.flags=r.r.u16(2);}if(r.type===4026)out[r.instance]=r.r.text(0,r.r.length&~1,1200).replace(/\0+$/,'');}}return out;};
    book.textHeaders=headers(book.documentRecords.filter(r=>r.type!==4057||r.instance!==4),envReport);
    book.noteHeaders=headers(book.documentRecords.filter(r=>r.type!==4057||r.instance!==3),envReport);
    const docExtensions=T.extensions(book.documentRecords,e=>envReport('format',e));book.textBlips=new Map();book.fonts10=new Map();
    const extendStyles=(records,map,report)=>{for(const rec of records)if(rec.type===4013||rec.type===4018)try{const levels=rec.type===4013?T.master9(rec).map(p=>({p})):T.master10(rec),styles=map.get(rec.instance)||[];levels.forEach((s,i)=>{styles[i]={p:{...styles[i]?.p,...s.p},r:{...styles[i]?.r,...s.r}};});map.set(rec.instance,styles);}catch(e){if(e.name!=='LegacyError')throw e;report('format',e);}};
    extendStyles(docExtensions,book.defaultStyles,envReport);
    for(const ext of docExtensions){
      if(ext.type===2040)for(const r of P.children(ext,e=>envReport('bullet',e)))if(r.type===2041)book.textBlips.set(r.instance,r);
      if(ext.type===2006)for(const r of P.children(ext,e=>envReport('format',e)))if(r.type===4023)try{r.r.check(0,68);book.fonts10.set(r.instance,{name:r.r.text(0,64,1200).split('\0')[0],charset:r.r.u8(64),pitchFamily:r.r.u8(67)});}catch(e){if(e.name!=='LegacyError')throw e;envReport('format',e);}
      if(ext.type===4020)try{book.defaultCF10=T.cf10(T.cursor(ext.r));}catch(e){if(e.name!=='LegacyError')throw e;envReport('format',e);}
    }
    const outlineExtensions=new Map();
    for(const ext of docExtensions)if([4014,4019,4021].includes(ext.type)){let key;for(const r of P.children(ext,e=>envReport('format',e))){if(r.type===4015)key=r.r.length>=4?r.r.u32(0)+':'+r.instance:undefined;else if([4012,4017,4022].includes(r.type)&&key){const records=outlineExtensions.get(key)||[];records.push(r);outlineExtensions.set(key,records);}}}
    for(const page of pages){
      const report=page.textReport=T.reporter(book,page);page.textStyles=new Map();page.styleCache=new Map();page.textShapes=[];page.textAll=[];page.textUsed=new Set();page.textHeaders=headers(page.children,report);readStyles(page.children,page.textStyles,report);
      page.textExtensions=T.extensions(page.children,e=>report('format',e));
      extendStyles(page.textExtensions,page.textStyles,report);
      const outlines=T.groups(page.listRecords),byIndex=new Map(outlines.map(g=>[g.index,g]));
      const parseText=records=>{const t=T.read(records,report);textBudget-=t.text.length;if(textBudget<0)B.fail('PPT_TEXT_LIMIT',N('Presentation text is too large'));return t;};
      const parsed=new Map();const outline=index=>{if(parsed.has(index))return parsed.get(index);const group=byIndex.get(index);if(!group)return null;const t=parseText(group.records);parsed.set(index,t);return t;};
      const node=(rec,depth=0)=>{
        take();if(depth>B.limits.depth)B.fail('PPT_TEXT_LIMIT',N('Presentation text nesting exceeds the limit'));
        const kids=P.children(rec,e=>report('drawing',e));
        if(rec.type!==0xf004)return kids.filter(c=>[0xf002,0xf003,0xf004].includes(c.type)).flatMap(c=>node(c,depth+1));
        const shape={properties:new Map(),children:[],records:kids};
        for(const c of kids)try{
          if(c.type===0xf00a){shape.id=c.r.u32(0);shape.flags=c.r.u32(4);shape.type=c.instance;}
          else if([0xf00b,0xf121,0xf122].includes(c.type))for(const [id,p]of A.properties(c))shape.properties.set(id,p);
          else if([0xf00f,0xf010].includes(c.type))shape.anchor=A.anchor(c,'ppt');
          else if(c.type===0xf009)shape.coord={left:c.r.i32(0),top:c.r.i32(4),right:c.r.i32(8),bottom:c.r.i32(12)};
        }catch(e){if(e.name!=='LegacyError')throw e;report('drawing',e);}
        if(shape.flags&0x400)return [];
        const client=kids.find(r=>r.type===0xf011),data=client?P.children(client,e=>report('format',e)):[];
        try{shape.ph=T.placeholder(data.find(r=>r.type===3011),data.find(r=>r.type===1056||r.type===3037));const original=data.find(r=>r.type===1055);if(original)shape.originalId=original.r.u32(0);}catch(e){if(e.name!=='LegacyError')throw e;report('format',e);}
        const box=kids.find(r=>r.type===0xf00d);let records=box?P.children(box,e=>report('text',e)):[];
        const ref=records.find(r=>r.type===3998);let text;
        if(ref){try{const index=ref.r.u32(0);text=outline(index);if(text){page.textUsed.add(index);shape.outline=index;}else report('reference',index);}catch(e){if(e.name!=='LegacyError')throw e;report('reference',e);}}
        if(!text&&records.some(r=>[3999,4000,4008].includes(r.type)))text=parseText(records);
        if(!text&&shape.ph)text=parseText([]);
        if(text){
          // Rulers in the referencing textbox override the outline's ruler.
          shape.text={...text};const ruler=records.find(r=>r.type===4006);if(ruler)try{shape.text.ruler=T.ruler(ruler);}catch(e){if(e.name!=='LegacyError')throw e;report('format',e);}
          for(const ext of T.extensions(data,e=>report('format',e)))try{if(ext.type===4012)shape.text.ext9=T.ext9(ext);if(ext.type===4017)shape.text.ext10=T.ext10(ext);}catch(e){if(e.name!=='LegacyError')throw e;report('format',e);}
          page.textAll.push(shape);
        }
        return [shape];
      };
      // Keep the group hierarchy so child anchors, flips and rotations remain
      // meaningful. The patriarch is the page coordinate system, not a shape.
      const tree=(rec,depth=0)=>{
        take();if(depth>B.limits.depth)B.fail('PPT_TEXT_LIMIT',N('Presentation text nesting exceeds the limit'));
        if(rec.type===0xf004)return node(rec,depth);
        const kids=P.children(rec,e=>report('drawing',e));
        if(rec.type===0xf003){const first=kids.find(c=>c.type===0xf004),group=first?node(first,depth+1)[0]:null,children=kids.filter(c=>c!==first&&[0xf003,0xf004].includes(c.type)).flatMap(c=>tree(c,depth+1));if(!group||group.flags&4)return children;group.children=children;return [group];}
        return kids.filter(c=>[0xf002,0xf003,0xf004].includes(c.type)).flatMap(c=>tree(c,depth+1));
      };
      for(const drawing of page.children.filter(r=>r.type===1036))page.textShapes.push(...tree(drawing));
      // Unreferenced outline text has no trustworthy placement. Retain its
      // loss explicitly; don't add guessed boxes or duplicate outline copies.
      for(const g of outlines)if(!page.textUsed.has(g.index)&&outline(g.index)?.text)report('text','Unreferenced outline '+g.index);
      const naked=page.children.filter(r=>r.type===4000||r.type===4008);if(naked.length)report('text','Text without a shape');
      // Extension outline blocks are keyed by source slide ID and text index.
      for(const shape of page.textAll)for(const ext of outlineExtensions.get(page.id+':'+shape.outline)||[])try{if(ext.type===4012)shape.text.ext9=T.ext9(ext);if(ext.type===4017)shape.text.ext10=T.ext10(ext);}catch(e){if(e.name!=='LegacyError')throw e;report('format',e);}
    }
    for(const page of pages)for(const shape of page.textAll){if(shape.ph&&shape.text.type===4&&!shape.text.text){shape.text.type=({title:0,ctrTitle:6,subTitle:5,body:page.kind.includes('notes')?2:1})[shape.ph.type]??4;}}
    for(const page of pages)for(const shape of page.textAll){
      const text=shape.text;let previous9=-1,index9=0,previous10=-1,index10=0;
      const apply=props=>{const n9=props.extIndex??0;if(n9!==previous9){index9=Math.floor(index9/16)*16+n9;if(previous9>=0&&n9<previous9)index9+=16;previous9=n9;}const extra=text.ext9?.[index9],n10=extra?.cf10Index??0;if(n10!==previous10){index10=Math.floor(index10/16)*16+n10;if(previous10>=0&&n10<previous10)index10+=16;previous10=n10;}return {p:extra,r:{...book.defaultCF10,...text.ext10?.[index10]}};};
      if(!text.runs.length)text.extra=apply({});for(const run of text.runs){run.extra=apply(run.props);Object.assign(run.props,run.extra.r);}
    }
    // OfficeArt frame/body properties inherit through hspMaster. A Boolean
    // property's upper half marks which lower bits override the master.
    const defaults=new Map(),shapes=new Map(),owners=new Map(),resolved=new Set(),active=new Set();
    const merge = A.mergeProperties;
    for(const dgg of book.documentRecords.filter(r=>r.type===1035))for(const root of P.children(dgg,e=>envReport('drawing',e)))if(root.type===0xf000)for(const rec of P.children(root,e=>envReport('drawing',e)))if([0xf00b,0xf122].includes(rec.type))try{merge(defaults,A.properties(rec));}catch(e){if(e.name!=='LegacyError')throw e;envReport('drawing',e);}
    const collect=(list,page)=>{for(const s of list){if(!shapes.has(s.id))shapes.set(s.id,s);owners.set(s,page);collect(s.children||[],page);}};for(const page of pages)collect(page.textShapes,page);
    const resolve=(s,depth=0)=>{if(resolved.has(s))return s.properties;if(active.has(s)||depth>B.limits.depth){owners.get(s).textReport('drawing','Cyclic master shape');return defaults;}active.add(s);const master=shapes.get(s.properties.get(769)?.value),props=new Map(master?resolve(master,depth+1):defaults);s.properties=merge(props,s.properties);active.delete(s);resolved.add(s);return s.properties;};
    for(const s of owners.keys())resolve(s);
    // A generated layout gets placeholder geometry from its first source
    // slide. The slide still owns all its own text/formatting. Stored OOXML
    // layouts are left byte-for-byte intact.
    const layoutSlides=new Map();book.slides.forEach((s,i)=>{s.textSlideNumber=book.firstSlideNumber+i;if(s.notes)s.notes.textSlideNumber=s.textSlideNumber;if(!layoutSlides.has(s.outputLayout))layoutSlides.set(s.outputLayout,s);});
    for(const layout of book.outputLayouts)layout.textPlaceholders=layoutSlides.get(layout)?.textAll.filter(s=>s.ph)||[];
    return book;
  };
  T.headersFor=(book,page)=>({...(/notes|handout/i.test(page.kind)?book.noteHeaders:book.textHeaders),...(page.kind==='notes'?book.notesMaster:page.parent)?.textHeaders,...page.textHeaders});
  T.missingHeaders=function(book,page){
    const h=T.headersFor(book,page),missing=[];
    for(const [bit,type,fields,string]of [[1,'dt',[4087,4088,4117],0],[8,'sldNum',[4056]],[16,'hdr',[4089],1],[32,'ftr',[4090],2]]){
      if(page.headerExpected?!page.headerExpected.includes(type):!(h.flags&bit)||[16,32].includes(bit)&&!h[string])continue;
      const found=page.textAll.some(s=>s.textWritten&&s.text.text&&(s.ph?.type===type||s.text.fields.some(f=>fields.includes(f.type))||h[string]&&s.text.text===h[string]));
      if(!found)missing.push(type);
    }return missing;
  };
})(typeof window!=='undefined'?window:globalThis);
