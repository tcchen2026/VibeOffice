// Independent binary fixtures: all coordinates and IDs are explicit test inputs.
export const u16=n=>{const b=Buffer.alloc(2);b.writeUInt16LE(n);return b;};
export const u32=n=>{const b=Buffer.alloc(4);b.writeUInt32LE(n>>>0);return b;};
export const rec=(type,data=Buffer.alloc(0),instance=0,version=0)=>Buffer.concat([u16(instance*16+version),u16(type),u32(data.length),data]);
export const container=(type,children=[],instance=0)=>rec(type,Buffer.concat(children),instance,15);
export const scheme=(colors,instance=1)=>rec(2032,Buffer.concat(colors.map(s=>Buffer.from(s+'00','hex'))),instance);
export const slideAtom=(master=0,notes=0,flags=7,layout=1)=>rec(1007,Buffer.concat([u32(layout),Buffer.alloc(8),u32(master),u32(notes),u16(flags),u16(0)]));
export function presentation(options={}){
 const master={id:0x80000000,ref:2,record:container(1016,[slideAtom(),scheme(['FFFFFF','000000','808080','123456','BBE0E3','333399','009999','99CC00'])])};
 const pages=options.pages||{masters:[master],slides:[{id:257,ref:3,record:container(1006,[slideAtom(master.id,300),rec(4008,Buffer.from('First slide'))])},{id:256,ref:4,record:container(1006,[slideAtom(master.id,0,3,16)])}],notes:[{id:300,ref:5,record:container(1008,[rec(1009,Buffer.concat([u32(257),u16(7),u16(0)])),rec(4008,Buffer.from('Speaker notes'))])}]};
 const size=options.size||[5760,4320,4320,5760],da=Buffer.concat([...size.map(u32),u32(1),u32(1),u32(options.notesRef||0),u32(options.handoutRef||0),u16(5),u16(0),Buffer.from([0,0,1,1])]);
 const font=Buffer.alloc(68);font.write('Arial',0,'utf16le');font[64]=0;font[67]=34;
 const env=container(1010,[container(2005,[rec(4023,font),...(options.fontRecords||[])])]);
 const lists=['slides','masters','notes'].map((key,instance)=>container(4080,(pages[key]||[]).map(p=>rec(1011,Buffer.concat([u32(p.ref),u32(0),u32(0),u32(p.id),u32(0)]))),instance));
 const doc=container(1000,[rec(1001,da),env,...lists,...(options.documentRecords||[])]),objects=[{ref:1,record:doc},...Object.values(pages).flat(),...(options.extra||[])];
 let at=16;const entries=[];for(const o of objects){entries.push({ref:o.ref,offset:at});at+=o.record.length;}
 const ptrAt=at,ptr=rec(6002,Buffer.concat([...entries.map(e=>Buffer.concat([u32(0x100000|e.ref),u32(e.offset)])),...(options.stale?[Buffer.concat([u32(0x100000|900),u32(0xfffffff0)])]:[])]));at+=ptr.length;
 const editAt=at,edit=rec(4085,Buffer.concat([u32(0),u32(0),u32(0),u32(ptrAt),u32(1),u32(1000),u32(0)]));
 const stream=Buffer.concat([Buffer.alloc(16),...objects.map(x=>x.record),ptr,edit]);
 const user=Buffer.alloc(28);user.writeUInt32LE(options.encrypted?0xf3d1c4df:0xe391c05f,12);user.writeUInt32LE(editAt,16);
 const streams={'PowerPoint Document':stream,'Current User':user};
 const fake={entries:[],stream:name=>streams[name]||null};return {fake,stream,user,bytes:cfb(streams),entries,pages};
}
export function cfb(streams){
 const items=Object.entries(streams).map(([name,data])=>({name,data,n:Math.ceil(Math.max(4096,data.length)/512)}));let next=2;
 for(const i of items){i.start=next;next+=i.n;}
 if(next>=128||items.length>3)throw Error('Fixture too large');
 const b=Buffer.alloc((next+1)*512),put=(p,v)=>b.writeUInt32LE(v>>>0,p);
 b.set([0xd0,0xcf,0x11,0xe0,0xa1,0xb1,0x1a,0xe1]);b.writeUInt16LE(3,26);b.writeUInt16LE(0xfffe,28);b.writeUInt16LE(9,30);b.writeUInt16LE(6,32);put(44,1);put(48,1);put(56,4096);put(60,0xfffffffe);put(68,0xfffffffe);
 for(let i=0;i<109;i++)put(76+i*4,i?0xffffffff:0);for(let i=0;i<128;i++)put(512+i*4,0xffffffff);put(512,0xfffffffd);put(516,0xfffffffe);
 const entries=[{name:'Root Entry',type:5,start:0xfffffffe,size:0},...items.map(i=>({...i,type:2,size:Math.max(4096,i.data.length)}))];
 entries.forEach((e,i)=>{const p=1024+i*128;b.write(e.name,p,'utf16le');b.writeUInt16LE((e.name.length+1)*2,p+64);b[p+66]=e.type;b[p+67]=1;put(p+68,0xffffffff);put(p+72,i&&i+1<entries.length?i+1:0xffffffff);put(p+76,i?0xffffffff:1);put(p+116,e.start);put(p+120,e.size);});
 for(const i of items){for(let n=0;n<i.n;n++)put(512+(i.start+n)*4,n===i.n-1?0xfffffffe:i.start+n+1);b.set(i.data,(i.start+1)*512);}
 return b;
}
