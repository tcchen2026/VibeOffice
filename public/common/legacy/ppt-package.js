/* Dependencies of OOXML designs stored inside binary PowerPoint records.
 * MS-PPT 2.11.9–12 and ECMA-376 Part 2. Bounded byte parsing only; external
 * relationships are retained as relationships and never fetched. */
(function(root){
  'use strict';
  const B=root.L.legacy,P=B.ppt,N=s=>s;
  const bad=()=>B.fail('PPT_EMBEDDED_LINKS',N('A stored presentation design has missing or damaged linked parts'));
  const local=n=>n.name.split(':').at(-1),ns=n=>n.ns[n.name.includes(':')?n.name.split(':')[0]:''];
  const relName=name=>{const i=name.lastIndexOf('/');return name.slice(0,i+1)+'_rels/'+name.slice(i+1)+'.rels';};
  const resolve=(from,target)=>{
    if(/[\\\x00-\x20?#]/.test(target))bad();let decoded;try{decoded=decodeURI(target);}catch{bad();}
    const parts=decoded.startsWith('/')?[]:from.split('/').slice(0,-1);
    for(const part of decoded.split('/')){if(!part||part==='.')continue;if(part==='..'){if(!parts.length)bad();parts.pop();}else parts.push(part);}
    return parts.join('/');
  };
  P.designBundle=async function(bytes,zip,entry,xml){
    const types=new Map(),defaults=new Map(),parts=new Map(),pending=[entry.name];let total=0;
    if(zip.has('[Content_Types].xml'))B.checkXml(B.decode(await B.zipPart(bytes,zip.get('[Content_Types].xml')),65001),node=>{if(ns(node)!==B.ns.content)return;if(local(node)==='Override')types.set(node.attrs.PartName?.replace(/^\//,''),node.attrs.ContentType);if(local(node)==='Default')defaults.set(node.attrs.Extension?.toLowerCase(),node.attrs.ContentType);});
    for(let i=0;i<pending.length;i++){
      const name=pending[i];if(parts.has(name))continue;if(parts.size>=256)bad();const item=zip.get(name);if(!item)bad();
      const data=name===entry.name?xml:await B.zipPart(bytes,item,64*1024*1024);total+=typeof data==='string'?data.length*2:data.length;if(total>64*1024*1024)bad();
      const type=types.get(name)||defaults.get(name.split('.').at(-1).toLowerCase());if(name!==entry.name&&(!type||/[\x00-\x20]/.test(type)))bad();
      const rels=[],ids=new Set(),file=zip.get(relName(name));
      if(file)B.checkXml(B.decode(await B.zipPart(bytes,file),65001),node=>{
        if(ns(node)!==B.ns.rel||local(node)!=='Relationship')return;
        const a=node.attrs;if(!/^[A-Za-z_][\w.:-]*$/.test(a.Id)||ids.has(a.Id)||!/^https?:\/\//.test(a.Type)||typeof a.Target!=='string'||/[\x00-\x1f]/.test(a.Target))bad();ids.add(a.Id);
        const rel={id:a.Id,type:a.Type,target:a.Target,external:a.TargetMode==='External',parent:name===entry.name&&a.Type===B.ns.officeRel+'slideMaster'};
        if(rel.external){if(/^(?:javascript|vbscript|data):/i.test(rel.target.trim()))bad();}
        else if(!rel.parent){rel.part=resolve(name,rel.target);if(!zip.has(rel.part))bad();pending.push(rel.part);if(pending.length>4096)bad();}
        rels.push(rel);
      });
      if(name===entry.name||/\+xml$|\/xml$/.test(type||'')||/\.xml$/i.test(name))B.checkXml(typeof data==='string'?data:B.decode(data,65001),node=>{for(const [key,value]of Object.entries(node.attrs))if(key.includes(':')&&node.ns[key.split(':')[0]]===B.ns.officeRel.slice(0,-1)&&!ids.has(value)&&!(value===''&&key.endsWith(':id')&&ns(node)===B.ns.a&&['hlinkClick','hlinkMouseOver'].includes(local(node))))bad();});
      parts.set(name,{data,type,rels});
    }
    return {root:entry.name,parts};
  };
  P.importDesign=function(book,pkg,design,name,parents={}){
    const bundle=design?.bundle,names=new Map(bundle?[[bundle.root,name]]:[]),numbers=book.designPartNumbers||(book.designPartNumbers=new Map());
    if(bundle)for(const [source,part]of bundle.parts){
      if(source===bundle.root)continue;let ext=source.split('.').at(-1).toLowerCase();if(!/^[a-z0-9]+$/.test(ext))ext='bin';let folder='embeddings',stem='oleObject';
      if(part.type.startsWith('image/')){folder='media';stem='image';}
      else if(part.type.includes('drawingml.chart')){folder='charts';stem='chart';}
      else if(part.type.includes('theme')){folder='theme';stem=part.type.includes('Override')?'themeOverride':'theme';}
      else if(part.type.includes('diagram')){folder='diagrams';stem=part.type.includes('data')?'data':part.type.includes('layout')?'layout':part.type.includes('colors')?'colors':'quickStyle';}
      else if(part.type.startsWith('audio/')||part.type.startsWith('video/')){folder='media';stem=part.type.startsWith('audio/')?'audio':'video';}
      const key=folder+'/'+stem;let n=numbers.get(key)||0,destination;do{destination='ppt/'+key+(++n)+'.'+ext;}while(pkg.parts.has(destination)||[...names.values()].includes(destination));numbers.set(key,n);names.set(source,destination);
      pkg.add(destination,part.data,part.type);
    }
    const parentTypes=new Set();
    if(bundle)for(const [source,part]of bundle.parts)for(const rel of part.rels){const from=names.get(source),kind=rel.type.slice(B.ns.officeRel.length);let target;
      if(source===bundle.root&&Object.hasOwn(parents,kind)){target=parents[kind];parentTypes.add(kind);}
      else if(rel.parent)continue;
      else target=rel.external?rel.target:P.drawings.relative(from,names.get(rel.part));
      pkg.rel(from,rel.type,target,{id:rel.id,external:rel.external});
    }
    for(const [kind,target]of Object.entries(parents))if(!parentTypes.has(kind))pkg.rel(name,kind,target);
  };
})(typeof window!=='undefined'?window:globalThis);
