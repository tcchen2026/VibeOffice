/* Binary Word complex fields and per-story positions.
 * Derived from Apache POI hwpf/model/{FieldsTables,FieldsDocumentPart,
 * FieldDescriptor,FFData,NilPICFAndBinData}.java (Apache-2.0).
 * Derived from Apache OpenOffice sw/source/filter/ww8/ww8par3.cxx
 * (Apache-2.0): form-control anchors and checkbox sizing.
 * Derived from Open XML SDK data/schemas/schemas_openxmlformats_org_wordprocessingml_2006_main.json (MIT): form property order.
 * MS-DOC 2.8.25 Plcfld, 2.9.88 Fld, 2.9.110 grffldEnd. Field instructions
 * and cached results are preserved; this reader never evaluates them. */
(function(root){
  'use strict';
  const B=root.L.legacy,D=B.doc,N=s=>s,F=D.fields={},w=(...a)=>D.w(...a),v=(name,val)=>w(name,{val});
  function form(book,node,where){
    const end=node.separator??node.end,kind=({FORMTEXT:0,FORMCHECKBOX:1,FORMDROPDOWN:2})[/^\s*([A-Z]+)/i.exec(book.text.slice(node.start+1,Math.min(end,node.start+40)))?.[1].toUpperCase()];
    if(kind==null)return '';
    const cp=end-1;node.formMark=book.text.charCodeAt(cp)===1?cp:null;
    return D.optional(book,where,N('A form field lost its settings; its last displayed text was retained'),()=>{
      if(node.formMark==null)B.fail('DOC_FORM_DATA',N('The form field settings are missing'));
      const c=D.properties(book,cp,2,null,where).props;
      if(c.pictureOffset==null)B.fail('DOC_FORM_DATA',N('The form field settings are missing'));
      book.data??=new B.Reader(book.cfb.stream('Data')||new Uint8Array(),'Data');
      const offset=c.pictureOffset,size=book.data.u32(offset),header=book.data.u16(offset+4),r=book.data.reader(offset+header,size-header,'Form field');
      if(header<6||r.u32(0)!==0xffffffff||(r.u16(4)&3)!==kind)B.fail('DOC_FORM_DATA',N('The form field settings are damaged'));
      const bits=r.u16(4),result=bits>>>2&31;let p=10;
      const string=max=>{const n=r.u16(p);p+=2;let s=D.unicode(r,p,n*2);p+=n*2;r.check(p,2);p+=2;
        if(s.length>max){D.notice(book,'legacy-word-form-settings',N('Some form field settings were shortened'),where,true);s=s.slice(0,max);if(/[\uD800-\uDBFF]$/.test(s))s=s.slice(0,-1);}return s;};
      const name=string(65),text=kind===0?string(255):null,def=kind===0?0:r.u16(p);if(kind!==0)p+=2;
      const format=string(64),help=string(255),status=string(138),entry=string(33),exit=string(33);
      let out=v('name',name)+v('enabled',bits&512?0:1)+v('calcOnExit',bits&16384?1:0);
      if(entry)out+=v('entryMacro',entry);if(exit)out+=v('exitMacro',exit);
      if(help)out+=w('helpText',{type:bits&128?'text':'autoText',val:help});
      if(status)out+=w('statusText',{type:bits&256?'text':'autoText',val:status});
      if(kind===0){const type=['regular','number','date','currentDate','currentTime','calculated'][bits>>>11&7];
        out+=w('textInput',{},(type?v('type',type):'')+v('default',text)+(r.u16(6)?v('maxLength',D.bound(book,r.u16(6),1,32767,where)):'')+(format?v('format',format):''));
      }else if(kind===1){
        // AOO's fUseSize and Office-written forms use 1 for a fixed hps size.
        out+=w('checkBox',{},(bits&1024?v('size',D.bound(book,r.u16(8),2,3168,where)):w('sizeAuto'))+v('default',def===1?1:0)+v('checked',result===1?1:0));
      }else{
        const all=D.markup.strings(r.reader(p,r.length-p)),items=all.slice(0,25);
        if(all.length>25||items.some(x=>x.text.length>255))D.notice(book,'legacy-word-form-settings',N('Some drop-down choices were shortened or not kept'),where,true);
        out+=w('ddList',{},(result<items.length?v('result',result):'')+(def<items.length?v('default',def):'')+items.map(x=>v('listEntry',x.text.slice(0,255).replace(/[\uD800-\uDBFF]$/,''))).join(''));
      }
      book.fieldStats.formFields++;return w('ffData',{},out);
    },'');
  }
  F.prepare=function(book){
    book.fieldStarts=new Map();book.fieldStats={sourceFields:0,fields:0,formFields:0,malformedFields:0};
    for(const [story,id]of[['main',16],['headers',17],['footnotes',18],['comments',19],['endnotes',48],['textboxes',57],['headerTextboxes',59]]){
      const range=book.storyRanges[story],where=story==='main'?'Document':story,metadata=new Map();
      if(book.pair(id).length)D.optional(book,where,N('Some field settings could not be read'),()=>{
        const r=D.table(book,id,'Field settings'),plc=D.plc(r,2);
        for(let i=0;i<plc.n;i++){
          const cp=range.start+plc.cp[i],at=plc.at(i),marker=r.u8(at)&31;
          if(cp>=range.end||book.text.charCodeAt(cp)!==marker||metadata.has(cp))continue;
          metadata.set(cp,{marker,flags:r.u8(at+1)});
        }
      },null);
      const stack=[];
      // XE/TC/RD/TA/PRIVATE fields intentionally have no PLC entry. The text
      // delimiters also recover readable fields after optional PLC damage.
      for(let cp=range.start;cp<range.end;cp++){
        const marker=book.text.charCodeAt(cp);
        if(marker===19){
          if(stack.length>=B.limits.depth)B.fail('DOC_FIELD_DEPTH',N('Fields in this document are nested too deeply'));
          const node={start:cp,separator:null,end:null,flags:0,type:metadata.get(cp)?.flags};
          stack.push(node);book.fieldStarts.set(cp,node);book.fieldStats.sourceFields++;
        }else if(marker===20&&stack.length){
          const node=stack.at(-1);if(node.separator==null)node.separator=cp;else node.damaged=true;
        }else if(marker===21&&stack.length){
          const node=stack.pop();node.end=cp;node.flags=metadata.get(cp)?.flags||0;
        }
      }
    }
  };
  F.begin=function(book,range,cp,where,parents){
    const node=book.fieldStarts?.get(cp),keep=!!node&&node.end!=null&&node.end<range.end&&!node.damaged&&parents.every(f=>f.keep||f.result);
    if(keep){book.fieldStats.fields++;node.formXML=form(book,node,where);}
    else{if(book.fieldStats)book.fieldStats.malformedFields++;D.notice(book,'legacy-word-fields',N('A damaged field was replaced with its last displayed text'),where,true);}
    return {node,keep,result:false};
  };
})(typeof window!=='undefined'?window:globalThis);
