/* Binary Word positions, formatting pages and stories (MS-DOC §§2.8–2.9).
 * Derived from Apache POI hwpf/model/{CHPFormattedDiskPage,PAPFormattedDiskPage,
 * PAPX,SectionTable,NotesTables}.java (Apache-2.0).
 * Derived from Apache OpenOffice sw/source/filter/ww8/ww8par.cxx (Apache-2.0).
 * CPs count UTF-16 units; FKP boundaries are physical byte offsets, including
 * in documents with reused or out-of-order pieces from incremental saves. */
(function(root){
  'use strict';
  const B=root.L.legacy,D=B.doc,N=s=>s;
  D.notice=function(book,id,what,where,notify=false,detail){
    const key=id+'\0'+where;book.noticeKeys??=new Map();
    if(book.noticeKeys.has(key)){const old=book.noticeKeys.get(key);old.count=(old.count||1)+1;return;}
    const loss=B.loss(id,what,where,'drop',where,notify);if(detail)loss.detail=detail;
    book.noticeKeys.set(key,loss);book.losses.push(loss);
  };
  D.optional=function(book,where,what,read,fallback,notify=true){
    try{return read();}catch(e){if(!['LegacyError','CFBError'].includes(e.name))throw e;D.notice(book,'legacy-word-damaged',what,where,notify,e.code);return fallback;}
  };
  D.plc=function(r,size){
    if(r.length<4||(r.length-4)%(size+4))B.fail('DOC_PLC',N('A document position table is damaged'));
    const n=(r.length-4)/(size+4),cp=[];
    if(n>B.limits.records)B.fail('DOC_PLC',N('A document position table is too large'));
    for(let i=0;i<=n;i++){const value=r.u32(i*4);if(i&&value<cp[i-1])B.fail('DOC_PLC',N('A document position table is out of order'));cp.push(value);}
    return {r,n,cp,at:i=>(n+1)*4+i*size};
  };
  D.table=function(book,id,label){const pair=book.pair(id);return book.tableReader.reader(pair.offset,pair.length,label);};
  D.locate=function(ranges,value){let lo=0,hi=ranges.length;while(lo<hi){const mid=(lo+hi)>>>1;if(ranges[mid].start<=value)lo=mid+1;else hi=mid;}return lo-1;};
  D.pieceAt=function(book,cp){const p=book.pieces[D.locate(book.pieces,cp)];return p&&cp<p.end?p:null;};
  D.at=function(book,ranges,cp){
    const piece=D.pieceAt(book,cp);if(!piece)return {};
    const step=piece.compressed?1:2,fc=piece.offset+(cp-piece.start)*step,index=D.locate(ranges,fc),run=ranges[index];
    return {piece,run:run&&fc<run.end?run:null,end:Math.min(piece.end,run&&fc<run.end?cp+Math.ceil((run.end-fc)/step):index+1<ranges.length?cp+Math.ceil((ranges[index+1].start-fc)/step):piece.end)};
  };
  function expand(book,ops,seen=new Set()){
    const out=[];
    for(const op of ops){
      if(op.code!==0x6645&&op.code!==0x6646){out.push(op);continue;}
      const offset=D.operand(op).u32(0);if(seen.has(offset)||seen.size>=B.limits.depth)B.fail('DOC_PAPX_CYCLE',N('Paragraph formatting contains a loop'));
      seen.add(offset);book.data??=new B.Reader(book.cfb.stream('Data')||new Uint8Array(),'Data');
      const n=book.data.u16(offset),inner=D.sprms(book.data.reader(offset+2,n,'Huge PAPX'));
      out.push(...expand(book,inner,seen));seen.delete(offset);
    }
    return out;
  }
  function pages(book,kind){
    const table=D.table(book,kind===1?13:12,kind===1?'PlcfBtePapx':'PlcfBteChpx');if(!table.length)return [];
    const plc=D.plc(table,4),seen=new Set(),runs=[];
    for(let i=0;i<plc.n;i++){
      const page=table.u32(plc.at(i))&0x3fffff;if(seen.has(page))continue;seen.add(page);
      D.optional(book,'Formatting page '+page,N('Some text formatting could not be read'),()=>{
        const f=book.r.reader(page*512,512,'FKP'),count=f.u8(511),stride=kind===1?13:1,base=(count+1)*4;
        if(base+count*stride>511)B.fail('DOC_FKP',N('A text formatting page is damaged'));
        const cache=new Map();
        for(let j=0;j<count;j++){
          const start=f.u32(j*4),end=f.u32((j+1)*4),at=f.u8(base+j*stride)*2;
          if(start>end)B.fail('DOC_FKP',N('Text formatting positions are out of order'));
          if(start===end)continue;
          let props=cache.get(at);
          if(!props){
            props=D.optional(book,'Formatting page '+page,N('Some text formatting could not be read'),()=>{
              if(!at)return {style:0,ops:[]};
              let p=at,n=f.u8(p++);if(kind===1){if(!n)n=f.u8(p++)*2;else n=n*2-1;}
              const r=f.reader(p,n,'FKP properties');return kind===1?{style:r.u16(0),ops:expand(book,D.sprms(r.reader(2,r.length-2)))}:{ops:D.sprms(r)};
            },{style:0,ops:[]});cache.set(at,props);
          }
          runs.push({start,end,...props});if(runs.length>B.limits.records)B.fail('DOC_FKP_LIMIT',N('The document has too many text formatting runs'));
        }
      },null);
    }
    runs.sort((a,b)=>a.start-b.start||a.end-b.end);return runs;
  }
  function sections(book){
    const r=D.table(book,6,'PlcfSed');if(!r.length)return [{start:0,end:book.stories.main,props:{}}];
    const plc=D.plc(r,12),out=[];
    for(let i=0;i<plc.n;i++){
      const offset=r.u32(plc.at(i)+2),where='Section '+(i+1);
      const props=D.optional(book,where,N('Some section formatting could not be read'),()=>{
        if(offset===0xffffffff)return {};
        return D.apply(D.sprms(book.r.reader(offset+2,book.r.u16(offset),'SEPX')),4);
      },{});
      if(plc.cp[i]<book.stories.main)out.push({start:plc.cp[i],end:Math.min(book.stories.main,plc.cp[i+1]),props,where});
    }
    if(!out.length)out.push({start:0,end:book.stories.main,props:{},where:'Section 1'});
    // Word tolerates a missing terminal PLC boundary. Keep the remaining text.
    out.at(-1).end=book.stories.main;return out;
  }
  function notes(book,name,referenceId,textId){
    const refs=D.table(book,referenceId,name+' references'),texts=D.table(book,textId,name+' text');
    if(!refs.length||!texts.length)return [];
    const rp=D.plc(refs,2),tp=D.plc(texts,0),result=[],story=book.storyRanges[name],claimed=new Set();
    for(let i=0;i<rp.n;i++){
      if(i>=tp.n||tp.cp[i+1]>story.end-story.start||rp.cp[i]>=book.stories.main){D.notice(book,'legacy-word-note',N('A note with damaged text positions was not kept'),name,true);continue;}
      const automatic=!!refs.u16(rp.at(i));let cp=rp.cp[i];
      // Some damaged PLCs point just after the automatic reference character.
      // Recover that marker without consuming the punctuation which follows it.
      if(automatic&&book.text[cp]!=='\x02'&&cp>0&&book.text[cp-1]==='\x02'&&!claimed.has(cp-1))cp--;
      claimed.add(cp);
      result.push({id:i+1,cp,automatic,start:story.start+tp.cp[i],end:story.start+tp.cp[i+1],where:(name==='footnotes'?'Footnote ':'Endnote ')+(i+1)});
    }
    return result;
  }
  function headers(book){
    const r=D.table(book,11,'PlcfHdd');if(!r.length)return [];
    const plc=D.plc(r,0),story=book.storyRanges.headers,out=[];
    for(let i=0;i<Math.min(plc.n,6+6*book.sections.length);i++){
      if(plc.cp[i+1]>story.end-story.start){D.notice(book,'legacy-word-header',N('A header or footer with damaged text positions was not kept'),'Headers and footers',true);continue;}
      let start=story.start+plc.cp[i],end=story.start+plc.cp[i+1];
      // Each nonempty subdocument has one final paragraph guard in addition to
      // its ordinary paragraph terminator. A zero-length range means inherit.
      if(end>start&&book.text[end-1]==='\r')end--;
      out.push({index:i,start,end,present:plc.cp[i+1]>plc.cp[i]});
    }
    return out;
  }
  D.structure=function(book,cfb,losses=[]){
    book.cfb=cfb;book.losses=losses;book.storyRanges={};let cp=0;
    for(const [name,n] of Object.entries(book.stories)){book.storyRanges[name]={start:cp,end:cp+n};cp+=n;}
    book.chpx=D.optional(book,'Document text',N('Character formatting could not be read'),()=>pages(book,2),[]);
    book.papx=D.optional(book,'Document text',N('Paragraph formatting could not be read'),()=>pages(book,1),[]);
    book.sections=D.optional(book,'Document',N('Section formatting could not be read'),()=>sections(book),[{start:0,end:book.stories.main,props:{}}]);
    book.headers=D.optional(book,'Headers and footers',N('Headers and footers could not be read'),()=>headers(book),[]);
    book.footnotes=D.optional(book,'Footnotes',N('Footnotes could not be read'),()=>notes(book,'footnotes',2,3),[]);
    book.endnotes=D.optional(book,'Endnotes',N('Endnotes could not be read'),()=>notes(book,'endnotes',46,47),[]);
    book.noteAt=new Map([...book.footnotes.map(n=>({...n,kind:'footnote'})),...book.endnotes.map(n=>({...n,kind:'endnote'}))].map(n=>[n.cp,n]));
    return book;
  };
})(typeof window!=='undefined'?window:globalThis);
