// Small binary Word fixtures. Offsets, piece order and FKP boundaries are explicit.
import {cfb,u16,u32} from './ppt-fixtures.mjs';
export {u16,u32};
export const sprm=(code,value)=>Buffer.concat([u16(code),Buffer.isBuffer(value)?value:code>>>13===0||code>>>13===1?Buffer.from([value]):code>>>13===3?u32(value):u16(value)]);
export const variable=(code,bytes)=>Buffer.concat([u16(code),Buffer.from([bytes.length]),bytes]);
const concat=parts=>Buffer.concat(parts);
export function style(name,kind=1,p=[],c=[],base=4095,sti=4094){
  const header=concat([u16(sti),u16(base*16+kind),u16(kind===1?2:1),u16(0),u16(0)]),nameBytes=concat([u16(name.length),Buffer.from(name+'\0','utf16le')]);
  const ops=kind===1?[concat([u16(0),...p]),concat(c)]:[concat(c)];
  const content=concat([header,nameBytes,...ops.map(b=>concat([u16(b.length),b,b.length%2?Buffer.alloc(1):Buffer.alloc(0)]))]);
  return concat([u16(content.length),content]);
}
export function doc(options={}){
  const stories={main:options.text??'Hello world\r',footnotes:'',headers:'',macros:'',comments:'',endnotes:'',textboxes:'',headerTextboxes:'',...options.stories},text=Object.values(stories).join('');
  const word=Buffer.alloc(12288),table=Buffer.alloc(8192),pairs=new Map();let tp=0;
  const add=(id,data)=>{pairs.set(id,[tp,data.length]);data.copy(table,tp);tp+=data.length;};
  const pieces=options.pieces||[{text,compressed:false,offset:2048}];let cp=0;
  for(const piece of pieces){piece.start=cp;cp+=piece.text.length;piece.end=cp;Buffer.from(piece.text,piece.compressed?'latin1':'utf16le').copy(word,piece.offset);}
  const plc=concat([...pieces.map(p=>u32(p.start)),u32(cp),...pieces.map(p=>concat([u16(0),u32(p.compressed?(p.offset*2)|0x40000000:p.offset),u16(p.prm||0)]))]);
  add(33,concat([...(options.prcs||[]).map(p=>concat([Buffer.from([1]),u16(p.length),p])),Buffer.from([2]),u32(plc.length),plc]));
  const fonts=options.fonts||['Times New Roman','Arial'];
  add(15,concat([u16(fonts.length),u16(0),...fonts.map(name=>{const f=Buffer.alloc(40);f[0]=40+(name.length+1)*2-1;f[1]=0x12;f.writeUInt16LE(400,2);return concat([f,Buffer.from(name+'\0','utf16le')]);})]));
  const styles=options.styles||[style('Normal',1,[],[sprm(0x4a43,24)],4095,0),style('Heading 1',1,[sprm(0x2406,1),sprm(0x2640,0)],[sprm(0x0835,1)],0,1)];
  add(1,concat([u16(18),u16(styles.length),u16(10),Buffer.alloc(14),...styles]));
  function fkp(id,page,runs,paragraph){
    const f=word.subarray(page*512,(page+1)*512),count=runs.length;f[511]=count;let pos=500;
    runs.forEach((run,i)=>{f.writeUInt32LE(run.start,i*4);f.writeUInt32LE(run.end,(i+1)*4);const bytes=paragraph?concat([u16(run.style||0),...(run.ops||[])]):concat(run.ops||[]);
      if(paragraph){const n=Math.ceil((bytes.length+1)/2);pos=(pos-1-bytes.length-1)&~1;f[pos]=n;bytes.copy(f,pos+1);}else{pos=(pos-1-bytes.length)&~1;f[pos]=bytes.length;bytes.copy(f,pos+1);}
      f[(count+1)*4+i*(paragraph?13:1)]=pos/2;
    });
    add(id,concat([u32(runs[0].start),u32(runs.at(-1).end),u32(page)]));
  }
  const byFC=[...pieces].sort((a,b)=>a.offset-b.offset),defaultRuns=byFC.map(p=>({start:p.offset,end:p.offset+p.text.length*(p.compressed?1:2),ops:[]}));
  fkp(12,8,options.chpx||defaultRuns,false);fkp(13,9,options.papx||defaultRuns,true);
  const sections=options.sections||[{end:stories.main.length,ops:[]}],sed=[];let previous=0;
  for(let i=0;i<sections.length;i++){const s=sections[i],bytes=concat(s.ops),offset=1024+i*128;u16(bytes.length).copy(word,offset);bytes.copy(word,offset+2);sed.push({start:previous,end:s.end,offset});previous=s.end;}
  add(6,concat([...sed.map(s=>u32(s.start)),u32(sed.at(-1).end),...sed.map(s=>concat([u16(0),u32(s.offset),u16(0),u32(0xffffffff)]))]));
  const dop=options.dop||Buffer.alloc(616);if(!options.dop){dop.writeUInt16LE(720,10);dop.writeUInt16LE(1,46);}add(31,dop);
  if(options.headers)add(11,concat(options.headers.map(u32)));
  for(const [name,rid,tid] of [['footnotes',2,3],['endnotes',46,47]])if(options[name]){const n=options[name];add(rid,concat([...n.map(x=>u32(x.cp)),u32(stories.main.length),...n.map(x=>u16(x.automatic===false?0:1))]));add(tid,concat([...n.map(x=>u32(x.start)),u32(stories[name].length)]));}
  for(const [id,data] of options.tables||[])add(id,data);
  word.writeUInt16LE(0xa5ec,0);word.writeUInt16LE(options.version??0xc1,2);word.writeUInt16LE(1033,6);word.writeUInt16LE(0x1204|(options.template?1:0)|(options.encrypted?0x100:0),10);
  word.writeUInt32LE(2048,24);word.writeUInt32LE(2048+text.length*2,28);
  word.writeUInt16LE(14,32);word.writeUInt16LE(22,62);let i=0;for(const name of Object.keys(stories))word.writeUInt32LE(stories[name].length,64+(3+i++)*4);
  word.writeUInt16LE(93,152);for(const [id,[offset,size]] of pairs){word.writeUInt32LE(offset,154+id*8);word.writeUInt32LE(size,158+id*8);}
  options.mutate?.(word,table,pairs);
  return cfb({'WordDocument':word,'1Table':table,...(options.data?{Data:options.data}: {})});
}
