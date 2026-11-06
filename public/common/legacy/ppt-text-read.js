/* Binary PowerPoint text records (MS-PPT §2.9).
 * Derived from Apache POI hslf/record/{StyleTextPropAtom,TxMasterStyleAtom,
 * TextRulerAtom,TextSpecInfoRun,StyleTextProp9Atom}.java and
 * hslf/model/textproperties/{TextPropCollection,TextPFException9}.java (Apache-2.0).
 * Derived from Apache OpenOffice filter/source/msfilter/svdfppt.cxx (Apache-2.0).
 * Counts are UTF-16 code units, including the implicit final paragraph break. */
(function(root){
  'use strict';
  const B=root.L.legacy,P=B.ppt,T=P.text={},N=s=>s;
  T.cursor=r=>({r,p:0,u16(){const v=r.u16(this.p);this.p+=2;return v;},i16(){const v=r.i16(this.p);this.p+=2;return v;},u32(){const v=r.u32(this.p);this.p+=4;return v;},skip(n){r.check(this.p,n);this.p+=n;}});
  const tabs=c=>{const n=c.u16();c.r.check(c.p,n*4);return Array.from({length:n},()=>({position:c.i16(),alignment:c.u16()}));};
  T.pf=function(c){
    const mask=c.u32(),p={};
    if(mask&15){const flags=c.u16();for(const [bit,key] of [[1,'bullet'],[2,'bulletHasFont'],[4,'bulletHasColor'],[8,'bulletHasSize']])if(mask&bit)p[key]=!!(flags&bit);}
    for(const [bit,key,kind] of [[128,'bulletChar','u16'],[16,'bulletFont','u16'],[64,'bulletSize','i16'],[32,'bulletColor','u32'],[2048,'align','u16'],[4096,'lineSpacing','i16'],[8192,'spaceBefore','i16'],[16384,'spaceAfter','i16'],[256,'margin','i16'],[1024,'indent','i16'],[32768,'defaultTab','i16']])if(mask&bit)p[key]=c[kind]();
    if(mask&0x100000)p.tabs=tabs(c);
    if(mask&0x10000)p.fontAlign=c.u16();
    if(mask&0xe0000){const flags=c.u16();for(const [bit,key,n] of [[0x20000,'charWrap',1],[0x40000,'wordWrap',2],[0x80000,'overflow',4]])if(mask&bit)p[key]=!!(flags&n);}
    if(mask&0x200000)p.rtl=!!c.u16();
    return p;
  };
  T.cf=function(c){
    const mask=c.u32(),p={};
    if(mask&65535){const flags=c.u16();for(const [bit,key] of [[1,'bold'],[2,'italic'],[4,'underline'],[16,'shadow'],[128,'kumimoji'],[256,'strike'],[512,'emboss']])if(mask&bit)p[key]=!!(flags&bit);if(mask&0x3c00)p.extIndex=(flags>>>10)&15;}
    for(const [bit,key,kind] of [[0x10000,'font','u16'],[0x200000,'eaFont','u16'],[0x400000,'ansiFont','u16'],[0x800000,'symbolFont','u16'],[0x20000,'size','u16'],[0x40000,'color','u32'],[0x80000,'baseline','i16']])if(mask&bit)p[key]=c[kind]();
    return p;
  };
  T.si=function(c){
    const mask=c.u32(),p={};
    if(mask&1)p.spelling=c.u16();if(mask&2)p.language=c.u16();if(mask&4)p.altLanguage=c.u16();
    if(mask&64)p.bidi=!!c.u16();if(mask&32)p.ext10=c.u32();if(mask&512)c.skip(c.u32()*4);
    return p;
  };
  T.pf9=function(c){const mask=c.u32(),p={};if(mask&0x800000)p.blip=c.u16();if(mask&0x2000000)p.autoNumber=!!c.u16();if(mask&0x1000000){p.numberScheme=c.u16();p.startNumber=c.u16();}if(mask&0x4000000)c.skip(4);return p;};
  T.cf9=c=>c.u32()&0x100000?c.u32()&15:0;
  T.ext9=function(rec){const c=T.cursor(rec.r),out=[];while(c.p<c.r.length){const p=T.pf9(c),cf10Index=T.cf9(c),si=T.si(c);out.push({...p,...si,cf10Index});if(out.length>B.limits.records)B.fail('PPT_TEXT_LIMIT',N('The presentation has too many text styles'));}return out;};
  // MS-PPT §§2.9.18, 2.9.39, 2.9.69. These indices use FontCollection10,
  // which is a separate table from the original FontCollection.
  T.cf10=function(c){const mask=c.u32(),p={};if(mask&0x1000000)p.eaFont10=c.u16();if(mask&0x2000000)p.csFont10=c.u16();if(mask&0x4000000)c.skip(4);return p;};
  T.ext10=function(rec){const c=T.cursor(rec.r),out=[];while(c.p<c.r.length)out.push(T.cf10(c));return out;};
  T.master10=function(rec){const c=T.cursor(rec.r),count=c.u16(),out=[];if(count>9)B.fail('PPT_TEXT_STYLE',N('A master text style has too many levels'));for(let i=0;i<count;i++)out.push({r:T.cf10(c)});return out;};
  T.ruler=function(rec){
    const c=T.cursor(rec.r),mask=c.u32(),out={levels:[]};if(mask&2)out.count=c.u16();if(mask&1)out.defaultTab=c.i16();if(mask&4)out.tabs=tabs(c);
    for(let i=0;i<5;i++){const p={};if(mask&(8<<i))p.margin=c.i16();if(mask&(256<<i))p.indent=c.i16();out.levels.push(p);}return out;
  };
  T.master=function(rec){
    const c=T.cursor(rec.r),count=c.u16(),out=[];
    if(count>9)B.fail('PPT_TEXT_STYLE',N('A master text style has too many levels'));
    for(let i=0;i<count;i++){const level=rec.instance>=5?c.u16():i,p=T.pf(c),r=T.cf(c);if(level<9)out[level]={p,r};}
    return out;
  };
  T.master9=function(rec){const c=T.cursor(rec.r),count=c.u16(),out=[];if(count>9)B.fail('PPT_TEXT_STYLE',N('A master text style has too many levels'));for(let i=0;i<count;i++){out.push(T.pf9(c));T.cf9(c);}return out;};
  T.groups=function(records){const groups=[];let group;for(const rec of records){if(rec.type===3999){group={index:groups.length,records:[]};groups.push(group);}if(group)group.records.push(rec);}return groups;};
  T.read=function(records,report){
    const atom=records.find(r=>r.type===4000||r.type===4008),header=records.find(r=>r.type===3999);
    const out={text:atom?(atom.type===4000?atom.r.text(0,atom.r.length&~1,1200):B.latin16(atom.r.bytes)):'',type:4,paragraphs:[],runs:[],special:[],fields:[]};
    const attempt=(rec,fn)=>{try{return fn();}catch(e){if(e.name!=='LegacyError')throw e;report('format',e.code+' in record '+rec.type);}};
    if(header)attempt(header,()=>{const n=header.r.u32(0);out.type=n<=8?n:4;});
    const count=out.text.length+1;
    const runs=(c,kind,target)=>{let pos=0;while(pos<count&&c.p<c.r.length){const n=c.u32();if(!n)B.fail('PPT_TEXT_RUN',N('A text formatting run has an invalid length'));const p=kind==='p'?{level:Math.min(8,c.u16()),...T.pf(c)}:kind==='r'?T.cf(c):T.si(c);const end=Math.min(count,pos+n);target.push({start:pos,end,props:p});pos=end;}if(pos<count)B.fail('PPT_TEXT_RUN',N('Text formatting ends before the text'));};
    for(const rec of records){
      if(rec.type===4001)attempt(rec,()=>{const c=T.cursor(rec.r);runs(c,'p',out.paragraphs);runs(c,'r',out.runs);});
      else if(rec.type===4002&&!records.some(r=>r.type===4001))attempt(rec,()=>{const c=T.cursor(rec.r);let pos=0;while(pos<count&&c.p<c.r.length){const n=c.u32(),level=Math.min(8,c.u16());if(!n)B.fail('PPT_TEXT_RUN',N('A text formatting run has an invalid length'));out.paragraphs.push({start:pos,end:Math.min(count,pos+n),props:{level}});pos+=n;}});
      else if(rec.type===4006)out.ruler=attempt(rec,()=>T.ruler(rec));
      else if(rec.type===4010)attempt(rec,()=>runs(T.cursor(rec.r),'s',out.special));
      else if(rec.type===4012)out.ext9=attempt(rec,()=>T.ext9(rec));
      else if([4056,4087,4088,4089,4090,4117].includes(rec.type))attempt(rec,()=>{const at=rec.r.u32(0);if(at<out.text.length){const field={at,type:rec.type};if(rec.type===4087)field.format=rec.r.u8(4);if(rec.type===4117)field.formatString=rec.r.text(4,Math.min(128,rec.r.length-4)&~1,1200).split('\0')[0];out.fields.push(field);}});
    }
    out.interactiveRecords=records.filter(r=>r.type===4082||r.type===4063);
    out.fields.sort((a,b)=>a.at-b.at);return out;
  };
  // The ninth-era extensions are inside a tagged binary blob, not ordinary
  // OfficeArt container payloads. Never interpret an arbitrary binary tag.
  T.extensions=function(records,report){
    const found=[],pending=records.map(rec=>({rec,depth:0}));let count=0;
    while(pending.length){const {rec,depth}=pending.pop();if(++count>B.limits.records||depth>B.limits.depth)B.fail('PPT_TEXT_LIMIT',N('Presentation text nesting exceeds the limit'));
      if(rec.type===5002){const children=P.children(rec,report),tag=children.find(r=>r.type===4026),blob=children.find(r=>r.type===5003);if(tag&&blob&&/^___PPT(?:9|10|11|12)$/.test(tag.r.text(0,tag.r.length,1200).replace(/\0+$/,'')))found.push(...P.children(blob,report));}
      else if(rec.type===5000||rec.type===2000)for(const child of P.children(rec,report))pending.push({rec:child,depth:depth+1});
    }return found;
  };
})(typeof window!=='undefined'?window:globalThis);
