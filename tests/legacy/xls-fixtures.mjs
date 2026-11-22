import assert from 'node:assert/strict';
const words=(...values)=>{const b=Buffer.alloc(values.length*2);values.forEach((v,i)=>b.writeUInt16LE(v,i*2));return b;};
const dword=n=>{const b=Buffer.alloc(4);b.writeUInt32LE(n);return b;};
const rec=(id,b=Buffer.alloc(0))=>Buffer.concat([words(id,b.length),b]);
const str=s=>Buffer.concat([words(s.length),Buffer.from([1]),Buffer.from(s,'utf16le')]);
const bof=type=>rec(0x809,words(0x600,type,0,0,0,0,0,0));
const bound=(name,offset,state=0)=>rec(0x85,Buffer.concat([dword(offset),Buffer.from([state,0,name.length,1]),Buffer.from(name,'utf16le')]));
function workbook(globals,sheets){
  const start=[bof(5),...globals,...sheets.map(s=>bound(s.name,0,s.state)),rec(10)].reduce((n,b)=>n+b.length,0);let at=start;
  const bodies=sheets.map(s=>Buffer.concat([bof(16),...s.records,rec(10)]));
  return Buffer.concat([bof(5),...globals,...sheets.map((s,i)=>{const b=bound(s.name,at,s.state);at+=bodies[i].length;return b;}),rec(10),...bodies]);
}
// Independent test-only CFB fixture, one regular stream; production has no CFB writer.
function cfb(data){
  const n=Math.ceil(Math.max(4096,data.length)/512),b=Buffer.alloc((n+3)*512),put=(p,v)=>b.writeUInt32LE(v>>>0,p);
  assert.ok(n<126);b.set([0xd0,0xcf,0x11,0xe0,0xa1,0xb1,0x1a,0xe1]);b.writeUInt16LE(3,26);b.writeUInt16LE(0xfffe,28);b.writeUInt16LE(9,30);b.writeUInt16LE(6,32);put(44,1);put(48,1);put(56,4096);put(60,0xfffffffe);put(68,0xfffffffe);
  for(let i=0;i<109;i++)put(76+i*4,i?0xffffffff:0);for(let i=0;i<128;i++)put(512+i*4,0xffffffff);
  put(512,0xfffffffd);put(516,0xfffffffe);for(let i=0;i<n;i++)put(512+(i+2)*4,i===n-1?0xfffffffe:i+3);
  for(const [i,name,type]of [[0,'Root Entry',5],[1,'Workbook',2]]){const p=1024+i*128;b.write(name,p,'utf16le');b.writeUInt16LE((name.length+1)*2,p+64);b[p+66]=type;b[p+67]=1;put(p+68,0xffffffff);put(p+72,0xffffffff);put(p+76,i?0xffffffff:1);put(p+116,i?2:0xfffffffe);put(p+120,i?Math.max(4096,data.length):0);}
  b.set(data,1536);return b;
}

export {words,dword,rec,str,bof,workbook,cfb};
