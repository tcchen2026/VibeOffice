/* BIFF formatting extensions, independently implemented from Microsoft MS-XLS:
 * XFExt 2.4.355, ExtProp 2.5.108, FullColorExt 2.5.155, GradStop 2.5.156,
 * XFPropGradient 2.5.286, ColorTheme 2.5.49 and Theme 2.4.326. */
(function(root){
  'use strict';
  const B=root.L.legacy,X=B.xls,N=s=>s,el=B.el;
  X.theme=async function(book,losses){
    const records=book.globals,i=records.findIndex(r=>r.id===0x896);if(i<0)return null;
    try{
      const first=records[i].r,chunks=[first.sub(16,first.length-16)];
      for(let j=i+1;records[j]?.id===0x812;j++){const r=records[j].r;chunks.push(r.sub(12,r.length-12));}
      // Default-theme records can omit their embedded ZIP; indexed BIFF colors
      // and the stored FONT names remain valid in that case.
      const size=chunks.reduce((n,b)=>n+b.length,0);if(!size)return null;if(size>4*1024*1024)B.fail('THEME_LIMIT',N('The workbook theme is too large'));
      const bytes=new Uint8Array(size);let p=0;for(const b of chunks){bytes.set(b,p);p+=b.length;}
      const zip=await root.L.zip.read(bytes),entry=[...zip.values()].find(e=>/(?:^|\/)theme\d*\.xml$/.test(e.name));
      const data=await B.zipPart(bytes,entry),xml=B.decode(data,65001),tree=B.checkXml(xml);
      const prefix=tree.name.includes(':')?tree.name.split(':')[0]:'';
      if(tree.name.split(':').at(-1)!=='theme'||tree.ns[prefix]!==B.ns.a)B.fail('THEME_XML',N('The workbook theme could not be read'));
      if([...zip.keys()].some(name=>name.endsWith('/_rels/'+entry.name.split('/').at(-1)+'.rels')))B.fail('THEME_MEDIA',N('The workbook theme contains linked parts'));
      return data;
    }catch(e){losses.push(B.loss('legacy-theme',N('The workbook theme was not kept'),'Workbook theme','drop'));return null;}
  };
  X.extColor=function(r,p,themeAvailable){
    const type=r.u16(p),tint=r.i16(p+2),value=r.u32(p+4);r.check(p,16);
    const c=X.colorValue(type,value,themeAvailable);if(c&&tint)c.tint=Math.max(-1,tint/32767);return c;
  };
  X.colorValue=function(type,value,themeAvailable){
    if(type===0)return {auto:1};if(type===1)return {indexed:value<=65?value:64};
    if(type===2)return {rgb:[value>>>24,value&255,value>>>8&255,value>>>16&255].map(n=>n.toString(16).padStart(2,'0')).join('').toUpperCase()};
    // BIFF orders dark/light pairs in the opposite order to SpreadsheetML.
    if(type===3&&value<12&&themeAvailable)return {theme:value<4?value^1:value};
    return null;
  };
  X.styleExtensions=function(book,xfs,losses){
    for(const rec of book.globals){const r=rec.r;
      try{
        if(rec.id===0x87d){
          const xf=xfs[r.u16(14)],count=r.u16(18);if(!xf)continue;let p=20;
          for(let i=0;i<count;i++){
            const type=r.u16(p),size=r.u16(p+2);if(size<4)B.fail('XF_EXTENSION',N('A cell format extension is damaged'));const data=r.reader(p+4,size-4);p+=size;
            if([4,5,7,8,9,10,11,13].includes(type)){
              const color=X.extColor(data,0,!!book.theme);if(!color)continue;
              const key={4:'foreground',5:'background',7:'top',8:'bottom',9:'left',10:'right',11:'diagonal',13:'fontColor'}[type];
              if(['top','bottom','left','right','diagonal'].includes(key))xf.border[key][1]=color;else xf[key]=color;
            }else if(type===14){const value=data.length===1?data.u8(0):data.u16(0);xf.scheme=book.theme?['none','major','minor'][value]:null;}
            else if(type===15)xf.align.indent=Math.min(250,data.u16(0));
            else if(type===6){
              const attrs=data.u32(0)?{type:'path',left:data.f64(12),right:data.f64(20),top:data.f64(28),bottom:data.f64(36)}:{type:'linear',degree:data.f64(4)};
              for(const [key,value]of Object.entries(attrs))if(key!=='type'&&(!Number.isFinite(value)||key!=='degree'&&(value<0||value>1)))B.fail('XF_GRADIENT',N('A cell gradient is damaged'));
              const n=data.u32(44),stops=[];if(n>256)B.fail('XF_GRADIENT',N('A cell gradient has too many stops'));data.check(48,n*22);
              for(let j=0;j<n;j++){const at=48+j*22,position=data.f64(at+6),tint=data.f64(at+14),color=X.colorValue(data.u16(at),data.u32(at+2),!!book.theme);if(!Number.isFinite(position)||position<0||position>1||!Number.isFinite(tint)||Math.abs(tint)>1||!color)B.fail('XF_GRADIENT',N('A cell gradient stop is damaged'));stops.push({position,color:{...color,...(tint?{tint}:{})}});}
              xf.gradient=el('gradientFill',attrs,stops.sort((a,b)=>a.position-b.position).map(s=>el('stop',{position:s.position},el('color',s.color))).join(''));
            }
          }
        }
      }catch(e){if(e.name!=='LegacyError')throw e;losses.push(B.loss('legacy-xls-format',N('Some workbook formatting could not be read'),rec.offset,'drop','Workbook formatting'));}
    }
  };
})(typeof window!=='undefined'?window:globalThis);
