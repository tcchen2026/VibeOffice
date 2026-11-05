/* Bounded XML validation for parts passed through from binary Office files.
 * The suite's general XML reader remains common/xml.js. This checks syntax,
 * namespaces and depth before retaining source bytes; it builds no element tree. */
(function(root){
  'use strict';
  const B=root.L.legacy,N=s=>s;
  const valid=s=>!/[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff]|[\ud800-\udbff](?![\udc00-\udfff])|(?<![\ud800-\udbff])[\udc00-\udfff]/.test(s);
  B.checkXml=function(text,visit){
    if(text.length>B.limits.text||!valid(text))B.fail('XML_TEXT',N('An embedded XML part contains invalid text'));
    const bad=()=>B.fail('XML_SYNTAX',N('An embedded XML part is damaged'));
    const decode=s=>s.replace(/&([^;]*);|&/g,(whole,entity)=>{
      if(!entity)bad();const named={lt:'<',gt:'>',amp:'&',quot:'"',apos:"'"};if(Object.hasOwn(named,entity))return named[entity];
      if(!/^#(?:[0-9]+|x[0-9a-fA-F]+)$/.test(entity))bad();const n=entity[1]==='x'?parseInt(entity.slice(2),16):Number(entity.slice(1));
      if(!Number.isInteger(n)||n>0x10ffff||n<0||n>=0xd800&&n<=0xdfff)bad();const ch=String.fromCodePoint(n);if(!valid(ch))bad();return ch;
    });
    const name=/[A-Za-z_\u0080-\ufffd][\w.\-:\u0080-\ufffd]*/y,space=/\s*/y;
    let p=0,tree=null,count=0;const stack=[];
    const skip=()=>{space.lastIndex=p;space.exec(text);p=space.lastIndex;};
    const readName=()=>{name.lastIndex=p;const m=name.exec(text);if(!m)bad();p=name.lastIndex;const parts=m[0].split(':');if(parts.length>2||parts.some(s=>!s))bad();return m[0];};
    const append=s=>{if(!stack.length&&s.trim())bad();if(visit?.text)visit.text(s,stack.at(-1));};
    while(p<text.length){
      if(text[p]!=='<'){const end=text.indexOf('<',p),q=end<0?text.length:end,s=text.slice(p,q);if(s.includes(']]>'))bad();append(decode(s));p=q;continue;}
      if(text.startsWith('<!--',p)){const end=text.indexOf('-->',p+4);if(end<0||text.slice(p+4,end).includes('--'))bad();p=end+3;continue;}
      if(text.startsWith('<?',p)){const end=text.indexOf('?>',p+2);if(end<0)bad();if(!/^<\?xml\s/.test(text.slice(p,end))||tree||stack.length)bad();p=end+2;continue;}
      if(text.startsWith('<![CDATA[',p)){const end=text.indexOf(']]>',p+9);if(end<0||!stack.length)bad();append(text.slice(p+9,end));p=end+3;continue;}
      if(text.startsWith('</',p)){const closeStart=p;p+=2;const tag=readName();skip();const node=stack.pop();if(text[p++]!=='>'||node?.name!==tag)bad();if(visit?.end)visit.end(node,closeStart,p);continue;}
      const start=p;p++;const tag=readName(),attrs=Object.create(null),parent=stack.at(-1);let empty=false;
      for(;;){const start=p;skip();if(text[p]==='>'){p++;break;}if(text.slice(p,p+2)==='/>'){empty=true;p+=2;break;}if(p===start)bad();const key=readName();if(Object.hasOwn(attrs,key))bad();skip();if(text[p++]!=='=')bad();skip();const quote=text[p++];if(quote!=="'"&&quote!=='"')bad();const end=text.indexOf(quote,p);if(end<0||text.slice(p,end).includes('<'))bad();attrs[key]=decode(text.slice(p,end));p=end+1;}
      const ns=Object.assign(Object.create(null),parent?.ns||{xml:'http://www.w3.org/XML/1998/namespace'});
      for(const [key,value]of Object.entries(attrs)){if(key==='xmlns')ns['']=value;else if(key.startsWith('xmlns:')){if(key==='xmlns:xmlns'||!value||key==='xmlns:xml'&&value!=='http://www.w3.org/XML/1998/namespace')bad();ns[key.slice(6)]=value;}}
      for(const key of [tag,...Object.keys(attrs)])if(key.includes(':')&&!key.startsWith('xmlns:')&&!ns[key.split(':')[0]])bad();
      const expanded=new Set();for(const key of Object.keys(attrs)){if(key==='xmlns'||key.startsWith('xmlns:'))continue;const parts=key.split(':'),full=(parts.length===2?ns[parts[0]]:'')+'|'+parts.at(-1);if(expanded.has(full))bad();expanded.add(full);}
      const node={name:tag,attrs,ns,start,openEnd:p};if(!parent){if(tree)bad();tree=node;}
      if(++count>B.limits.records||stack.length>=B.limits.depth)B.fail('XML_LIMIT',N('An embedded XML part is too complex'));
      if(visit)visit(node,stack.length);
      if(!empty)stack.push(node);else if(visit?.end)visit.end(node,p,p);
    }
    if(stack.length||!tree)bad();return tree;
  };
  // ZIP metadata comes from the shared reader; bounded decompression is needed
  // here because its general-purpose entry.bytes() accepts a growing output.
  B.zipPart=async function(bytes,entry,limit=4*1024*1024){
    if(!entry||!Number.isSafeInteger(entry.usize)||entry.usize<0||entry.usize>limit)B.fail('ZIP_PART_LIMIT',N('An embedded part is too large'));
    const r=new B.Reader(bytes,'Embedded ZIP'),p=entry.loff;if(r.u32(p)!==0x04034b50)B.fail('ZIP_PART',N('An embedded part is damaged'));
    const start=p+30+r.u16(p+26)+r.u16(p+28),raw=r.sub(start,entry.csize);let result;
    if(entry.method===0)result=raw;
    else if(entry.method===8){
      const chunks=[];let n=0,reader;
      try{reader=new Blob([raw]).stream().pipeThrough(new DecompressionStream('deflate-raw')).getReader();
        for(;;){const {done,value}=await reader.read();if(done)break;n+=value.length;if(n>entry.usize||n>limit){await reader.cancel();B.fail('ZIP_PART_LIMIT',N('An embedded part exceeds its declared size'));}chunks.push(value);}
      }catch(e){if(e.name==='LegacyError')throw e;B.fail('ZIP_COMPRESSION',N('An embedded part could not be decompressed'));}finally{reader?.releaseLock();}
      result=new Uint8Array(n);let at=0;for(const c of chunks){result.set(c,at);at+=c.length;}
    }else B.fail('ZIP_COMPRESSION',N('An embedded part uses unsupported compression'));
    if(result.length!==entry.usize||entry.crc!=null&&root.L.zip.crc32(result)!==entry.crc)B.fail('ZIP_PART',N('An embedded part is damaged'));
    return result;
  };
})(typeof window!=='undefined'?window:globalThis);
