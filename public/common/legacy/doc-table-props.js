/* Word table row/cell properties, MS-DOC 2.6.3, TDefTableOperand and TC80.
 * Derived from Apache OpenOffice sw/source/filter/ww8/ww8par2.cxx
 * (Apache-2.0): row operands, legacy horizontal merge flags and border defaults.
 * Derived from Apache POI hwpf/sprm/TableSprmUncompressor.java and
 * hwpf/usermodel/TableCellDescriptor.java (Apache-2.0).
 * Each row is bounded to Word's 63 cells; missing TC80 tails use defaults. */
(function(root){
  'use strict';
  const B=root.L.legacy,D=B.doc,N=s=>s,T=D.tables={},sides=['top','left','bottom','right','insideH','insideV'];
  T.width=(r,p=0)=>({type:r.u8(p)&15,value:r.i16(p+1)});
  const cell=()=>({borders:{},padding:{}});
  function definition(p,r){
    const n=r.u8(0);if(n>63)B.fail('DOC_TABLE_COLUMNS',N('A table row has too many cells'));
    r.check(1,2*(n+1));p.edges=Array.from({length:n+1},(_,i)=>r.i16(1+2*i));
    // A later TDefTable can provide only a prefix of the TC80s. With the same
    // column count, AOO ReadDef retains the others and separately set padding
    // and shading. A changed column count starts a new set of cell defaults.
    const cells=p.cells.length===n?p.cells:Array.from({length:n},cell);
    for(let i=0;i<n;i++){
      const at=1+2*(n+1)+i*20,c=cells[i];
      if(at+20<=r.length){const f=r.u16(at);Object.assign(c,{hmerge:f&3,flow:f>>>2&7,vmerge:f>>>5&3,align:f>>>7&3,width:{type:f>>>9&7,value:r.u16(at+2)},fit:!!(f&4096),noWrap:!!(f&8192),hideMark:!!(f&16384)});
        c.borders={};
        for(let j=0;j<4;j++)if(r.u32(at+4+j*4))c.borders[sides[j]]=D.border(r.reader(at+4+j*4,4));
      }
    }
    p.cells=cells;
  }
  const subset=(p,r)=>p.cells.slice(Math.min(63,r.u8(0)),Math.min(63,r.u8(1)));
  T.apply=function(book,ops,where){
    const p={cells:[],edges:[0],borders:{},padding:{top:0,left:108,bottom:0,right:108},gap:0,autoFit:false};
    for(const op of ops){
      if(op.kind!==5)continue;const r=D.operand(op),code=op.code;
      // A malformed optional visual operand loses only that property.
      D.optional(book,where,N('Some table formatting could not be read'),()=>{
        if(code===0xd608){definition(p,r);return;}
        if(code===0x5400||code===0x548a){p.jc=r.u16(0);p.logical=code===0x548a;return;}
        if(code===0x9601){p.left=r.i16(0);return;}
        if(code===0x9602){p.gap=Math.max(0,r.i16(0));return;}
        if(code===0x3403||code===0x3466){p[code===0x3466?'cantSplit':'cantSplitOld']=!!r.u8(0);return;}
        if(code===0x3404){p.header=!!r.u8(0);return;}
        if(code===0x9407){p.height=r.i16(0);return;}
        if(code===0x560b||code===0x5664){p.rtl=p.rtl||!!r.u16(0);return;}
        if(code===0x3615){p.autoFit=!!r.u8(0);return;}
        if(code===0x3619){p.keepFollow=!!r.u8(0);return;}
        if(code===0x3465){p.noOverlap=!!r.u8(0);return;}
        if(code===0x360d){p.position=r.u8(0);return;}
        const position={0x940e:'x',0x940f:'y',0x9410:'textLeft',0x9411:'textTop',0x941e:'textRight',0x941f:'textBottom'}[code];if(position){p[position]=r.i16(0);return;}
        const width={0xf614:'width',0xf617:'before',0xf618:'after',0xf661:'indent'}[code];if(width){p[width]=T.width(r);return;}
        if(code===0x740a){p.look=r.u16(2);return;}
        if(code===0xd605||code===0xd613){const n=code===0xd605?4:8;for(let i=0;i<6;i++)p.borders[sides[i]]=D.border(r.reader(i*n,n));return;}
        if(code>=0xd61a&&code<=0xd61d){const side=sides[code-0xd61a];for(let i=0;i<Math.min(p.cells.length,r.length/4);i++){const b=p.cells[i].borders[side];if(b)p.cells[i].borders[side]={...b,color:D.color(r.u32(i*4))};}return;}
        if(code===0xd609||[0xd612,0xd616,0xd60c,0xd670,0xd671,0xd672].includes(code)){
          const size=code===0xd609?2:10,start=code===0xd616||code===0xd671?22:code===0xd60c||code===0xd672?44:0;
          for(let i=0;i<Math.min(p.cells.length-start,Math.floor(r.length/size));i++)p.cells[start+i].shading=D.shading(r.reader(i*size,size));return;
        }
        if(code===0xd660){p.shading=D.shading(r);return;}
        if(code===0xd620||code===0xd62f){const mask=r.u8(2),border=D.border(r.reader(3,code===0xd620?4:8));for(const c of subset(p,r))for(let i=0;i<6;i++)if(mask&(1<<i))c.borders[sides[i]]=border;return;}
        if(code===0xd662){for(let i=0;i<Math.min(p.cells.length,Math.floor(r.length/4));i++)for(let j=0;j<4;j++){const c=p.cells[i],side=sides[j];c.borders[side]={...(c.borders[side]||{color:'auto',size:4,space:0}),type:r.u8(i*4+j)};}return;}
        if(code===0xd632||code===0xd634||code===0xd633){
          const mask=r.u8(2),size=r.u8(3)===0?0:r.u16(4);
          if(code===0xd633){p.spacing=size;return;}
          const targets=code===0xd634?[p]:subset(p,r);for(const c of targets)for(let i=0;i<4;i++)if(mask&(1<<i))c.padding[sides[i]]=size;return;
        }
        if(code===0xd635){const width=T.width(r,2);for(const c of subset(p,r))c.width=width;return;}
        if(code===0x7629||code===0xf629){for(const c of subset(p,r))c.flow=op.n===3?r.u8(2):r.u16(2);return;}
        if(code===0xd62b){const c=p.cells[r.u8(0)];if(c)c.vmerge=r.u8(1);return;}
        if(code===0xd62c){for(const c of subset(p,r))c.align=r.u8(2);return;}
        if(code===0xd62d||code===0xd62e){const list=subset(p,r),shading=D.shading(r.reader(2,10));list.forEach((c,i)=>{if(code===0xd62d||!(i&1))c.shading=shading;});return;}
        if(code===0xf636||code===0xd639||code===0xd642){const key=code===0xf636?'fit':code===0xd639?'noWrap':'hideMark';for(const c of subset(p,r))c[key]=!!r.u8(2);return;}
        if(code===0x5624||code===0x5625){subset(p,r).forEach((c,i)=>c.hmerge=code===0x5625?0:i?2:1);return;}
        if(code===0x7621){const first=Math.min(r.u8(0),p.cells.length),count=r.u8(1),width=r.u16(2);if(count+p.cells.length>63)B.fail('DOC_TABLE_COLUMNS',N('A table row has too many cells'));
          const at=p.edges[first];p.edges.splice(first+1,0,...Array.from({length:count},(_,i)=>at+(i+1)*width));for(let i=first+count+1;i<p.edges.length;i++)p.edges[i]+=count*width;p.cells.splice(first,0,...Array.from({length:count},cell));return;
        }
        if(code===0x5622){const a=Math.min(r.u8(0),p.cells.length),z=Math.min(r.u8(1),p.cells.length);if(z>a){const delta=p.edges[z]-p.edges[a];p.cells.splice(a,z-a);p.edges.splice(a+1,z-a);for(let i=a+1;i<p.edges.length;i++)p.edges[i]-=delta;}return;}
        if(code===0x7623){const a=Math.min(r.u8(0),p.cells.length),z=Math.min(r.u8(1),p.cells.length),width=r.u16(2);if(z>a){const delta=(z-a)*width-(p.edges[z]-p.edges[a]);for(let i=a+1;i<=z;i++)p.edges[i]=p.edges[a]+(i-a)*width;for(let i=z+1;i<p.edges.length;i++)p.edges[i]+=delta;}return;}
        if(code===0x3668){if(r.u8(0)){const copy={...p};delete copy.revisionBefore;p.revisionBefore=JSON.parse(JSON.stringify(copy));}else delete p.revisionBefore;return;}
        if(code===0xd667){p.revisionFormatting=!!r.u8(0);p.formatRevision={author:r.i16(1),date:r.u32(3)};return;}
        if(code===0x7479){p.rsid=r.u32(0);return;}
        if(code===0x7469)return; // XML group identity; no table-layout effect.
        D.notice(book,'legacy-word-table-format',N('Some table formatting was not kept'),where,true,'0x'+code.toString(16));
      },null);
    }
    if(p.left!=null){const delta=p.left-p.gap-p.edges[0];p.edges=p.edges.map(x=>x+delta);}
    for(let i=0;i<p.cells.length;i++)if(p.edges[i+1]<=p.edges[i])B.fail('DOC_TABLE_GRID',N('A table row has damaged cell widths'));
    return p;
  };
  T.properties=function(book,node,where){
    const {piece,run}=D.at(book,book.papx,node.z-1),chain=[],seen=new Set();let style=node.style;
    while(style&&!seen.has(style.index)&&chain.length<B.limits.depth){seen.add(style.index);chain.push(style);style=book.styles[style.basedOn];}
    return T.apply(book,[...chain.reverse().flatMap(s=>s.pOps),...(run?.ops||[]),...(piece?D.pieceSprms(book,piece):[])],where);
  };
})(typeof window!=='undefined'?window:globalThis);
