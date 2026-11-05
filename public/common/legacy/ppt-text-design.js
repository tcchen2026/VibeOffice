/* MS-PPT §§2.11.16–19: saved OOXML placeholder identities and text styles.
 * Self-contained fragments keep the original namespaces and style XML.
 * Linked fragments fall back locally to the native text properties. */
(function(root){
  'use strict';
  const B=root.L.legacy,P=B.ppt,T=P.text,N=s=>s;
  const local=n=>n.name.split(':').at(-1),namespace=n=>n.ns[n.name.includes(':')?n.name.split(':')[0]:''];
  const declarations=node=>Object.fromEntries(Object.entries(node.ns).filter(([key])=>key!=='xml').map(([key,value])=>[key?'xmlns:'+key:'xmlns',value]));
  T.embedded=async function(book){
    let total=0;book.textStats.storedTextStyles=0;
    for(const page of [...book.masters,book.notesMaster].filter(Boolean))for(const rec of page.children.filter(r=>r.type===1059||r.type===1063))try{
      if(rec.r.length>4*1024*1024)B.fail('PPT_TEXT_STYLE',N('An embedded text style is too large'));
      const zip=await root.L.zip.read(rec.r.bytes),entries=[...zip.values()].filter(e=>/(?:^|\/)slideMaster\d*\.xml$/.test(e.name));
      if(entries.length!==1)B.fail('PPT_TEXT_STYLE',N('An embedded text style could not be read'));
      const data=await B.zipPart(rec.r.bytes,entries[0]);total+=data.length;
      if(data.length>4*1024*1024||total>B.limits.text)B.fail('PPT_TEXT_STYLE',N('Embedded text styles are too large'));
      const xml=B.decode(data,65001);let styles,body,fragment,notes,linked=false;
      const visit=node=>{if(namespace(node)===B.ns.p){if(local(node)==='txStyles')styles=node;if(local(node)==='bodyStyle')body=node;}for(const key of Object.keys(node.attrs))if(key.includes(':')&&node.ns[key.split(':')[0]]===B.ns.officeRel.slice(0,-1))linked=true;};
      visit.end=(node,close)=>{if(node===styles)fragment=B.el('p:txStyles',{...declarations(node),'xmlns:p':B.ns.p},xml.slice(node.openEnd,close));if(node===body)notes=B.el('p:notesStyle',{...declarations(node),'xmlns:p':B.ns.p},xml.slice(node.openEnd,close));};
      B.checkXml(xml,visit);
      if(!fragment||linked)B.fail('PPT_TEXT_STYLE',N('An embedded text style contains unavailable content'));
      if(rec.type===1063){if(!notes)B.fail('PPT_TEXT_STYLE',N('An embedded notes style could not be read'));page.storedNotesStyle=notes;}
      else page.storedTextStyles=fragment;
      book.textStats.storedTextStyles++;
    }catch(e){if(!['LegacyError','Error','TypeError'].includes(e.name))throw e;page.textReport('format',e);}
    const layouts=new Map();
    for(const slide of book.slides){
      const layout=slide.outputLayout;if(!layout?.stored)continue;
      if(!layouts.has(layout)){
        const shapes=[],stack=[];
        const visit=node=>{if(namespace(node)!==B.ns.p)return;const name=local(node);if(name==='sp')stack.push({node});if(name==='cNvPr'&&stack.length)stack.at(-1).id=+node.attrs.id;if(name==='ph'&&stack.length)stack.at(-1).ph={...node.attrs,type:node.attrs.type||'obj',idx:+(node.attrs.idx||0)};};
        visit.end=node=>{if(stack.at(-1)?.node===node){const s=stack.pop();if(s.ph)shapes.push(s);}};
        B.checkXml(layout.stored.xml,visit);layouts.set(layout,shapes);
      }
      const candidates=layouts.get(layout);
      for(const shape of slide.textAll.filter(s=>s.ph?.idx==null&&s.ph)){
        const same=candidates.filter(s=>s.ph.type===shape.ph.type),match=same.find(s=>s.id===shape.originalId)||(same.length===1?same[0]:null);
        if(match)shape.ph.idx=match.ph.idx;
      }
    }
  };
})(typeof window!=='undefined'?window:globalThis);
