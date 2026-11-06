/* PowerPoint headers/footers and slide transitions.
 * Derived from Apache OpenOffice sd/source/filter/ppt/pptin.cxx (Apache-2.0).
 * Derived from Apache POI hslf/model/HeadersFooters.java (Apache-2.0).
 * MS-PPT 2.4.15, 2.6.6: visibility, fields, transition variants and timing. */
(function(root){
  'use strict';
  const B=root.L.legacy,P=B.ppt,T=P.text,D=P.drawings,el=B.el,F=P.features={};
  const headers=[{bit:1,type:'dt',field:4088,string:0},{bit:8,type:'sldNum',field:4056},{bit:16,type:'hdr',field:4089,string:1},{bit:32,type:'ftr',field:4090,string:2}];
  F.prepare=function(book){
    let modern=false;
    for(const m of book.masters)for(const rec of m.children.filter(r=>r.type===5000))P.walk(rec,r=>{if(r.type===4026&&r.r.length===16&&r.r.text(0,16,1200)==='___PPT12')modern=true;},()=>{});
    book.drawingStats.headerFields=0;
    for(const page of [...book.slides,...book.notes]){
      const h=T.headersFor(book,page),omit=page.kind==='slide'&&book.omitTitlePlaceholders&&page.layout===0;
      page.headerExpected=[];
      for(const spec of headers){
        const own=page.textAll.filter(s=>s.ph?.type===spec.type),placeholders=modern||h.flags==null,visible=placeholders?own.some(s=>s.text?.text&&(s.text.text!=='*'||s.text.fields.length)):!omit&&!!(h.flags&spec.bit)&&(spec.string==null||spec.type==='dt'&&h.flags&2||!!h[spec.string]);
        if(!visible){if(!placeholders)for(const s of own)s.headerHidden=true;continue;}
        page.headerExpected.push(spec.type);if(own.length)continue;
        let parent=page.kind==='notes'?book.notesMaster:page.parent,source;const seen=new Set();
        while(parent&&!seen.has(parent)){seen.add(parent);source=parent.textAll?.find(s=>s.ph?.type===spec.type&&s.anchor);if(source)break;parent=parent.parent;}
        if(!source)continue; // Missing geometry is reported, never guessed.
        const original=source.text||{},text={...original,text:'*',type:original.type??4,paragraphs:original.paragraphs?.slice(0,1)||[],runs:original.runs?.slice(0,1)||[],special:original.special?.slice(0,1)||[],fields:[{at:0,type:spec.field}],interactiveRecords:[]};
        const shape={...source,headerClone:true,properties:new Map(source.properties),ph:{...source.ph},children:[],records:[],text};
        page.textShapes.push(shape);page.textAll.push(shape);book.drawingStats.headerFields++;
      }
    }
  };
  F.headerFlags=function(book,page){
    const h=T.headersFor(book,page);if(h.flags==null)return '';
    return el('p:hf',{hdr:+!!(h.flags&16&&h[1]),ftr:+!!(h.flags&32&&h[2]),dt:+!!(h.flags&1&&(h.flags&2||h[0])),sldNum:+!!(h.flags&8)});
  };
  F.transition=function(book,page){
    let owner=page,rec;const seen=new Set();while(owner&&!seen.has(owner)){seen.add(owner);rec=owner.children.find(r=>r.type===1017);if(rec)break;owner=owner.parent;}
    if(!rec)return '';const r=rec.r;
    try{
      r.check(0,16);const direction=r.u8(8),type=r.u8(9),flags=r.u16(10),speed=r.u8(12),time=r.i32(0),attrs={spd:['slow','med',null][speed],advClick:flags&1?null:0,advTm:flags&1024?Math.max(0,Math.min(86399000,time)):null};
      const directions=['l','u','r','d','lu','ru','ld','rd'];let effect='';
      if(type===0)effect=direction===1?el('p:cut',{thruBlk:1}):'';
      else if(type===1)effect=el('p:random');
      else if(type===2)effect=el('p:blinds',{dir:direction?'horz':'vert'});
      else if(type===3)effect=el('p:checker',{dir:direction?'vert':'horz'});
      else if(type===4||type===7)effect=el(type===4?'p:cover':'p:pull',{dir:directions[direction]||'l'});
      else if(type===5)effect=el('p:dissolve');
      else if(type===6||type===23)effect=el('p:fade',{thruBlk:type===6?1:null});
      else if(type===8)effect=el('p:randomBar',{dir:direction?'vert':'horz'});
      else if(type===9)effect=el('p:strips',{dir:directions[direction>=4&&direction<=7?direction:4]});
      else if(type===10||type===20)effect=el(type===10?'p:wipe':'p:push',{dir:directions[direction<4?direction:0]});
      else if(type===11)effect=el('p:zoom',{dir:direction?'in':'out'});
      else if(type===13)effect=el('p:split',{orient:direction&2?'vert':'horz',dir:direction&1?'in':'out'});
      else if([17,18,19,22,27].includes(type))effect=el('p:'+({17:'diamond',18:'plus',19:'wedge',22:'newsflash',27:'circle'})[type]);
      else if(type===21)effect=el('p:comb',{dir:direction?'vert':'horz'});
      else if(type===26)effect=el('p:wheel',{spokes:[1,2,3,4,8].includes(direction)?direction:4});
      else if(type!==255)D.report(book,page,'transition',null,'Unknown effect '+type);
      if(flags&16)D.report(book,page,'sound',null,'Transition sound');
      if(flags&1024&&(time<0||time>86399000))D.report(book,page,'transition',null,'Invalid advance time');
      if(speed>2&&type!==0&&type!==255)D.report(book,page,'transition',null,'Invalid transition speed');
      page.transitionWritten=true;book.drawingStats.transitions++;
      return el('p:transition',attrs,effect+(flags&256?el('p:sndAc',{},el('p:endSnd')):''));
    }catch(e){if(e.name!=='LegacyError')throw e;D.report(book,page,'transition',null,e);return '';}
  };
})(typeof window!=='undefined'?window:globalThis);
