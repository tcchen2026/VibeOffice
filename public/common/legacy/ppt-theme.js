/* MS-PPT 2.11.9–12, 2.11.20 and 2.11.27: embedded OOXML themes/layouts.
 * MS-PPT 2.5.10–12: old eight-color schemes retain their semantic slots.
 * ECMA-376 DrawingML theme structure; Office-authored corpus examples are
 * recorded in tmp/legacy/D1-SOURCES.md. No DOM or external resources. */
(function(root){
  'use strict';
  const B=root.L.legacy,P=B.ppt,el=B.el,N=s=>s;
  P.colorMap={bg1:'lt1',tx1:'dk1',bg2:'lt2',tx2:'dk2',accent1:'accent1',accent2:'accent2',accent3:'accent3',accent4:'accent4',accent5:'accent5',accent6:'accent6',hlink:'hlink',folHlink:'folHlink'};
  P.schemeXml=function(scheme,name=''){
    const colors=[scheme[1],scheme[0],scheme[3],scheme[2],scheme[4],scheme[5],scheme[6],scheme[7],scheme[4],scheme[5],scheme[6],scheme[7]];
    return el('a:clrScheme',{name},['dk1','lt1','dk2','lt2','accent1','accent2','accent3','accent4','accent5','accent6','hlink','folHlink'].map((tag,i)=>el('a:'+tag,{},el('a:srgbClr',{val:colors[i]}))).join(''));
  };
  P.themeXml=function(page,book){
    const font=book.fonts.get(0)||book.fonts.values().next().value||{name:'Arial',charset:0,pitchFamily:0};
    const body=el('a:latin',{typeface:font.name,pitchFamily:font.pitchFamily>127?font.pitchFamily-256:font.pitchFamily,charset:font.charset>127?font.charset-256:font.charset})+el('a:ea',{typeface:''})+el('a:cs',{typeface:''});
    const solid=el('a:solidFill',{},el('a:schemeClr',{val:'phClr'}));
    const format=el('a:fmtScheme',{name:''},el('a:fillStyleLst',{},solid.repeat(3))+el('a:lnStyleLst',{},[9525,25400,38100].map(w=>el('a:ln',{w,cap:'flat',cmpd:'sng',algn:'ctr'},solid+el('a:prstDash',{val:'solid'})+el('a:miter',{lim:800000}))).join(''))+el('a:effectStyleLst',{},el('a:effectStyle',{},el('a:effectLst')).repeat(3))+el('a:bgFillStyleLst',{},solid.repeat(3)));
    const extra=page.schemeGallery?.length?el('a:extraClrSchemeLst',{},page.schemeGallery.map(s=>el('a:extraClrScheme',{},P.schemeXml(s)+el('a:clrMap',P.colorMap))).join('')):'';
    return B.xmlDeclaration+el('a:theme',{'xmlns:a':B.ns.a,name:page.name||''},el('a:themeElements',{},P.schemeXml(page.scheme)+el('a:fontScheme',{name:''},el('a:majorFont',{},body)+el('a:minorFont',{},body))+format)+el('a:objectDefaults')+extra);
  };
  P.roundTrip=async function(rec,kind){
    if(rec.r.length>4*1024*1024)B.fail('PPT_EMBEDDED_LIMIT',N('An embedded presentation design is too large'));
    const bytes=rec.r.bytes,zip=await root.L.zip.read(bytes),entries=[...zip.values()].filter(e=>kind==='layout'?/(?:^|\/)slideLayout\d*\.xml$/.test(e.name):/(?:^|\/)theme(?:Override)?\d*\.xml$/.test(e.name));
    if(entries.length!==1)B.fail('PPT_EMBEDDED_PART',N('An embedded presentation design could not be read'));
    const entry=entries[0],data=await B.zipPart(bytes,entry),xml=B.decode(data,65001),meta={};let linked=false;
    const tree=B.checkXml(xml,(node,depth)=>{
      if(depth===1&&node.name.split(':').at(-1)==='cSld')meta.name=node.attrs.name;
      for(const key of Object.keys(node.attrs))if(key.includes(':')&&node.ns[key.split(':')[0]]===B.ns.officeRel.slice(0,-1))linked=true;
    });
    const local=tree.name.split(':').at(-1),ns=tree.ns[tree.name.includes(':')?tree.name.split(':')[0]:''];
    if(kind==='layout'?local!=='sldLayout'||ns!==B.ns.p:!['theme','themeOverride'].includes(local)||ns!==B.ns.a)B.fail('PPT_EMBEDDED_PART',N('An embedded presentation design could not be read'));
    // Retain the bounded dependency graph with its original relationship IDs.
    const slash=entry.name.lastIndexOf('/'),rels=entry.name.slice(0,slash+1)+'_rels/'+entry.name.slice(slash+1)+'.rels';
    const bundle=linked||zip.has(rels)?await P.designBundle(bytes,zip,entry,xml):null;
    return {xml,root:local,attrs:tree.attrs,bundle,...meta};
  };
  P.prepareDesigns=async function(book,losses){
    const pages=[...book.masters,...book.slides,...book.notes,book.notesMaster,book.handoutMaster].filter(Boolean),originals=new Map();
    for(const page of pages){
      page.roundTripLayouts=[];page.colorMap={...P.colorMap};
      for(const rec of page.children){
        try{
          if(rec.type===1038)page.theme=await P.roundTrip(rec,'theme');
          else if(rec.type===1039){let mapping;B.checkXml(B.decode(rec.r.bytes,65001),node=>{const parts=node.name.split(':'),local=parts.at(-1),ns=node.ns[parts.length===2?parts[0]:''];if(['clrMap','overrideClrMapping'].includes(local)&&[B.ns.a,B.ns.p].includes(ns))mapping=node.attrs;});if(!mapping)continue;const map={};for(const key of Object.keys(P.colorMap)){const val=mapping[key];if(!['dk1','lt1','dk2','lt2','accent1','accent2','accent3','accent4','accent5','accent6','hlink','folHlink'].includes(val))B.fail('PPT_COLOR_MAP',N('A presentation color mapping could not be read'));map[key]=val;}page.colorMap=map;page.hasStoredColorMap=true;}
          else if(rec.type===1052)page.originalMasterId=rec.r.u32(0);
          else if(rec.type===1053)page.compositeMasterId=rec.r.u32(0);
          else if(rec.type===1058)page.originalLayout={master:rec.r.u32(0),id:rec.r.u32(4)};
          else if(rec.type===1054){const data=await P.roundTrip(rec,'layout');page.roundTripLayouts.push({...data,sourceId:rec.instance,source:page});}
        }catch(e){
          if(!['LegacyError','Error','TypeError'].includes(e.name))throw e;
          P.notice(losses,'legacy-presentation-design',N('A stored presentation design could not be kept'),page.where,rec.type===1054,e.code||e.message);
        }
      }
      if(page.originalMasterId!=null&&!originals.has(page.originalMasterId))originals.set(page.originalMasterId,page);
    }
    // Source master IDs are retained. Office allocates layout IDs from the same
    // high unsigned range, with no overlap; reserve all source IDs first.
    const used=new Set(book.masters.map(m=>m.id).filter(id=>id>=0x80000000)),assigned=new Set();let next=0x80000000;
    const id=preferred=>{if(preferred>=0x80000000&&!assigned.has(preferred)){assigned.add(preferred);return preferred;}while(used.has(next)||assigned.has(next))next++;if(next>0xffffffff)B.fail('PPT_LIMIT',N('The presentation has too many master layouts'));const n=next++;assigned.add(n);return n;};
    const masters=[],layouts=[],bySource=new Map(),rtLayouts=new Map();
    const masterFor=source=>{if(bySource.has(source))return bySource.get(source);const m={source,id:id(source?.id),layouts:[],layoutByKey:new Map(),index:masters.length+1};masters.push(m);bySource.set(source,m);return m;};
    for(const source of book.masters){
      const original=originals.get(source.compositeMasterId);
      if(source.record?.type===1006&&source.parent||original&&original!==source)continue;
      masterFor(source);
    }
    const layoutFor=(m,key,source,type='blank',stored=null)=>{
      let layout=m.layoutByKey.get(key);if(layout)return layout;
      layout={key,master:m,source,type,stored,id:id(),index:layouts.length+1};layouts.push(layout);m.layouts.push(layout);m.layoutByKey.set(key,layout);return layout;
    };
    for(const source of book.masters){
      const original=originals.get(source.compositeMasterId),base=original&&original!==source?original:source;
      if(source.record?.type===1006&&source.parent)continue;
      const m=masterFor(base);bySource.set(source,m);
      for(const stored of source.roundTripLayouts){const key='stored:'+stored.sourceId,layout=layoutFor(m,key,source,stored.attrs.type||'cust',stored);rtLayouts.set(m.index+':'+stored.sourceId,layout);}
    }
    const resolveMaster=source=>{const seen=new Set();while(source?.record?.type===1006&&source.parent){if(seen.has(source)||seen.size>=B.limits.depth)return masterFor(null);seen.add(source);source=source.parent;}return bySource.get(source)||masterFor(source);};
    for(const source of book.masters)if(source.record?.type===1006&&source.parent){const m=resolveMaster(source);source.titleLayout=layoutFor(m,'title:'+source.id,source,'title');}
    const layoutTypes={0:'title',1:'tx',2:'title',7:'titleOnly',8:'twoColTx',9:'twoRowTx',10:'txAndTwoObj',11:'twoObjAndTx',12:'objOverTx',13:'twoObjOverTx',14:'fourObj',15:'objOnly',16:'blank',17:'vertTitleAndTx',18:'vertTx'};
    for(const slide of book.slides){
      const source=slide.parent,m=resolveMaster(source);let layout=source?.titleLayout;
      if(!layout&&slide.originalLayout){const original=originals.get(slide.originalLayout.master),owner=original&&bySource.get(original);if(owner)layout=rtLayouts.get(owner.index+':'+slide.originalLayout.id);}
      if(!layout&&source?.compositeMasterId!=null&&source.roundTripLayouts.length===1)layout=rtLayouts.get(m.index+':'+source.roundTripLayouts[0].sourceId);
      if(!layout)layout=layoutFor(m,'native:'+slide.layout+':'+slide.placeholders.join(','),source,layoutTypes[slide.layout]||'cust');
      slide.outputLayout=layout;
    }
    if(!masters.length)masterFor(null);
    for(const m of masters)if(!m.layouts.length)layoutFor(m,'blank',m.source,'blank');
    book.outputMasters=masters;book.outputLayouts=layouts;return book;
  };
})(typeof window!=='undefined'?window:globalThis);
