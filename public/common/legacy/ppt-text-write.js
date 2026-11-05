/* PresentationML text, using the binary readers' resolved style properties.
 * Derived from Apache POI sl/usermodel/AutoNumberingScheme.java and
 * hslf/usermodel/HSLFTextShape.java (Apache-2.0).
 * MS-PPT §2.9 and ECMA-376 DrawingML text: paragraphs, runs and fields. */
(function(root){
  'use strict';
  const B=root.L.legacy,P=B.ppt,T=P.text,S=B.shapes,el=B.el,N=s=>s;
  const emu=v=>Math.round(v*1587.5),clamp=(v,a,b)=>Math.min(b,Math.max(a,v));
  const numbering=['alphaLcPeriod','alphaUcPeriod','arabicParenR','arabicPeriod','romanLcParenBoth','romanLcParenR','romanLcPeriod','romanUcPeriod','alphaLcParenBoth','alphaLcParenR','alphaUcParenBoth','alphaUcParenR','arabicParenBoth','arabicPlain','romanUcParenBoth','romanUcParenR','ea1ChsPlain','ea1ChsPeriod','circleNumDbPlain','circleNumWdWhitePlain','circleNumWdBlackPlain','ea1ChtPeriod','ea1ChtPlain','arabic1Minus','arabic2Minus','hebrew2Minus','ea1JpnKorPlain','ea1JpnKorPeriod','arabicDbPlain','arabicDbPeriod','thaiAlphaPeriod','thaiAlphaParenR','thaiAlphaParenBoth','thaiNumPeriod','thaiNumParenR','thaiNumParenBoth','hindiAlphaPeriod','hindiNumPeriod','ea1JpnChsDbPeriod','hindiNumParenR','hindiAlpha1Period'];
  const languages={1025:'ar-SA',1028:'zh-TW',1029:'cs-CZ',1030:'da-DK',1031:'de-DE',1032:'el-GR',1033:'en-US',1034:'es-ES',1035:'fi-FI',1036:'fr-FR',1037:'he-IL',1038:'hu-HU',1040:'it-IT',1041:'ja-JP',1042:'ko-KR',1043:'nl-NL',1044:'nb-NO',1045:'pl-PL',1046:'pt-BR',1048:'ro-RO',1049:'ru-RU',1050:'hr-HR',1051:'sk-SK',1053:'sv-SE',1054:'th-TH',1055:'tr-TR',1057:'id-ID',1058:'uk-UA',1060:'sl-SI',1061:'et-EE',1062:'lv-LV',1063:'lt-LT',1066:'vi-VN',1081:'hi-IN',2052:'zh-CN',2057:'en-GB',2070:'pt-PT',3076:'zh-HK',3082:'es-ES',4105:'en-CA',5129:'en-NZ'};
  const font=(book,index)=>book.fonts.get(index);
  const charset=n=>n>127?n-256:n;
  const fontXml=(tag,book,index,modern=false)=>{const f=modern?book.fonts10.get(index):font(book,index);return f?el(tag,{typeface:f.name,pitchFamily:charset(f.pitchFamily),charset:charset(f.charset)}):'';};
  T.color=function(value,page){const index=value>>>24;if(index===255)return '';const rgb=index<8?(page?.scheme||P.defaultScheme)[index]:[value&255,value>>>8&255,value>>>16&255].map(v=>v.toString(16).padStart(2,'0')).join('').toUpperCase();return rgb?el('a:srgbClr',{val:rgb}):'';};
  T.rPr=function(tag,r,book,page){
    const color=r.color==null?'':T.color(r.color,page),attrs={lang:languages[r.language],altLang:languages[r.altLanguage],sz:r.size?clamp(Math.round(r.size*100),100,400000):null,b:r.bold==null?null:+r.bold,i:r.italic==null?null:+r.italic,u:r.underline==null?null:r.underline?'sng':'none',strike:r.strike==null?null:r.strike?'sngStrike':'noStrike',baseline:r.baseline==null?null:clamp(r.baseline*1000,-100000,100000),kumimoji:r.kumimoji?1:null};
    const effect=r.shadow?el('a:effectLst',{},el('a:outerShdw',{blurRad:0,dist:38100,dir:2700000,algn:'ctr',rotWithShape:0},el('a:srgbClr',{val:page.scheme?.[2]||'808080'}))):'';
    const latin=font(book,r.font)?.charset===2?r.font:book.fonts.has(r.ansiFont)?r.ansiFont:r.font;
    return el(tag,attrs,(color?el('a:solidFill',{},color):'')+effect+fontXml('a:latin',book,latin)+fontXml('a:ea',book,book.fonts10.has(r.eaFont10)?r.eaFont10:r.eaFont,book.fonts10.has(r.eaFont10))+fontXml('a:cs',book,book.fonts10.has(r.csFont10)?r.csFont10:r.csFont,book.fonts10.has(r.csFont10))+fontXml('a:sym',book,r.symbolFont));
  };
  T.pPr=function(tag,p,r,book,page){
    const attrs={lvl:tag==='a:pPr'?p.level:null,marL:p.margin==null?null:clamp(emu(p.margin),0,51206400),indent:p.indent==null?null:clamp(emu(p.indent-(p.margin||0)),-51206400,51206400),algn:['l','ctr','r','just','dist','thaiDist','justLow'][p.align],defTabSz:p.defaultTab==null?null:clamp(emu(p.defaultTab),0,51206400),rtl:p.rtl==null?null:+p.rtl,fontAlgn:['base','t','ctr','b'][p.fontAlign],eaLnBrk:p.charWrap==null?null:+p.charWrap,latinLnBrk:p.wordWrap==null?null:+!p.wordWrap,hangingPunct:p.overflow==null?null:+p.overflow};
    let body='';
    for(const [key,name]of [['lineSpacing','lnSpc'],['spaceBefore','spcBef'],['spaceAfter','spcAft']])if(p[key]!=null){const v=p[key];body+=el('a:'+name,{},el(v<0?'a:spcPts':'a:spcPct',{val:v<0?Math.round(-v*12.5):v*1000}));}
    if(p.bullet){
      body+=p.bulletHasColor&&p.bulletColor!=null?el('a:buClr',{},T.color(p.bulletColor,page)||T.color(r.color??0x01000000,page)):el('a:buClrTx');
      body+=p.bulletHasSize&&p.bulletSize?el(p.bulletSize<0?'a:buSzPts':'a:buSzPct',{val:p.bulletSize<0?clamp(-p.bulletSize*100,100,400000):clamp(p.bulletSize*1000,25000,400000)}):el('a:buSzTx');
      const bf=font(book,p.bulletHasFont?p.bulletFont:r.font),symbol=bf?.charset===2;
      body+=bf?el('a:buFont',{typeface:bf.name,pitchFamily:charset(bf.pitchFamily),charset:charset(bf.charset)}):el('a:buFontTx');
      const c=p.bulletChar||0x2022,character=String.fromCharCode(symbol&&c<256?c|0xf000:c);
      const picture=p.blip!=null&&p.blip!==65535?book.bulletRel?.(p.blip):null;
      body+=picture?el('a:buBlip',{},el('a:blip',{'r:embed':picture})):p.autoNumber?el('a:buAutoNum',{type:numbering[p.numberScheme??3]||'arabicPeriod',startAt:clamp(p.startNumber||1,1,32767)}):el('a:buChar',{char:character});
      if(p.blip!=null&&p.blip!==65535&&!picture)page.textReport?.('bullet','Picture bullet '+p.blip);
    }else body+=el('a:buNone');
    if(p.tabs?.length)body+=el('a:tabLst',{},p.tabs.filter(t=>t.position>=0).map(t=>el('a:tab',{pos:emu(t.position),algn:['l','ctr','r','dec'][t.alignment]||'l'})).join(''));
    return el(tag,attrs,body+T.rPr('a:defRPr',r,book,page));
  };
  const rangeAt=(list,pos,cursor)=>{while(cursor.i<list.length&&list[cursor.i].end<=pos)cursor.i++;const r=list[cursor.i];return r&&r.start<=pos?r:null;};
  const typeForPh=ph=>({title:0,ctrTitle:6,subTitle:5,body:1})[ph?.type];
  T.dateText=function(n,language,date){
    const numeric={year:'numeric',month:'numeric',day:'numeric'},long={year:'numeric',month:'long',day:'numeric'},short={year:'numeric',month:'short',day:'numeric'},time={hour:'numeric',minute:'2-digit'};
    const formats=[numeric,{...long,weekday:'long'},long,long,{...short,year:'2-digit'},{year:'2-digit',month:'long'},{year:'2-digit',month:'short'},{...numeric,...time,hour12:true},{...numeric,...time,second:'2-digit',hour12:true},{...time,hour12:false},{...time,second:'2-digit',hour12:false},{...time,hour12:true},{...time,second:'2-digit',hour12:true}];
    return new Intl.DateTimeFormat(languages[language]||'en-US',formats[n]||numeric).format(date);
  };
  T.field=function(field,book,page,shape,r={}){
    const headers=T.headersFor(book,page);
    if(field.type===4056)return {type:'slidenum',text:String(page.textSlideNumber??book.firstSlideNumber)};
    if(field.type===4089||field.type===4090)return {text:headers[field.type===4089?1:2]||''};
    if(field.type===4088&&headers.flags!=null&&!(headers.flags&2))return {text:headers[0]||''};
    if(field.type===4117){const format=field.formatString||'';if(/^(?:'[^']*')+$/.test(format))return {text:format.replace(/'/g,'')};page.textReport?.('field',format);return {text:shape.text.text[field.at]||''};}
    const format=field.format??headers.format??0,n=format>=0&&format<=12?format:0;return {type:'datetime'+(n+1),text:T.dateText(n,r.language??r.altLanguage,book.textDate||(book.textDate=new Date()))};
  };
  const symbolText=(s,r,book)=>font(book,r.font)?.charset===2?s.replace(/[\x20-\xff]/g,c=>String.fromCharCode(c.charCodeAt(0)|0xf000)):s;
  T.body=function(shape,book,page,blank=false){
    const t=shape.text||{text:'',type:typeForPh(shape.ph)??4,paragraphs:[],runs:[],special:[],fields:[]},text=blank?'':t.text,parts=[];
    const pc={i:0},rc={i:0},sc={i:0},softPattern=/[\v\n]/g;let start=0,fieldIndex=0,paragraphs=0,nextSoft=softPattern.exec(text)?.index??text.length;
    do{
      const breakAt=text.indexOf('\r',start),end=breakAt<0?text.length:breakAt,localP=rangeAt(t.paragraphs,start,pc)?.props||{},level=localP.level||0,style=T.style(book,page,t.type,level),ruler=t.ruler||{},p={...style.p,...ruler.levels?.[level],...localP,level},baseR={...style.r,...book.defaultSI};
      if(ruler.defaultTab!=null)p.defaultTab=ruler.defaultTab;if(ruler.tabs)p.tabs=ruler.tabs;
      const first=rangeAt(t.runs,start,rc),firstRun=first?.props||t.extra?.r||{},ext=first?.extra?.p||t.extra?.p;if(ext)Object.assign(p,ext);
      const defaultR={...baseR,...firstRun};let runs='',pos=start;
      while(pos<end){
        const rr=rangeAt(t.runs,pos,rc),sr=rangeAt(t.special,pos,sc),r={...baseR,...t.extra?.r,...rr?.props,...sr?.props};
        while(fieldIndex<t.fields.length&&t.fields[fieldIndex].at<pos)fieldIndex++;
        const f=t.fields[fieldIndex];
        if(r.emboss)page.textReport?.('format','Embossed text');
        if(f?.at===pos){const data=T.field(f,book,page,shape,r),content=T.rPr('a:rPr',r,book,page)+el('a:t',{},B.xml(data.text));runs+=data.type?el('a:fld',{id:root.L.opc.guid(),type:data.type},content):el('a:r',{},content);book.textStats.fields++;pos++;fieldIndex++;continue;}
        if(text[pos]==='\v'||text[pos]==='\n'){runs+=el('a:br',{},T.rPr('a:rPr',r,book,page));pos++;continue;}
        while(nextSoft<pos)nextSoft=softPattern.exec(text)?.index??text.length;
        let until=Math.min(end,rr?.end??end,sr?.end??end,f?.at??end,nextSoft);
        // Bad styles may split a surrogate pair; its scalar must stay intact.
        if(until<end&&until>pos&&/[\ud800-\udbff]/.test(text[until-1])&&/[\udc00-\udfff]/.test(text[until]))until++;
        if(until<=pos)until=pos+1;
        runs+=el('a:r',{},T.rPr('a:rPr',r,book,page)+el('a:t',{},B.xml(symbolText(text.slice(pos,until),r,book))));pos=until;
      }
      const terminal={...baseR,...rangeAt(t.runs,end,rc)?.props,...rangeAt(t.special,end,sc)?.props};
      parts.push(el('a:p',{},T.pPr('a:pPr',p,defaultR,book,page)+runs+T.rPr('a:endParaRPr',terminal,book,page)));paragraphs++;start=end+1;
      if(breakAt<0)break;
    }while(start<=text.length);
    if(!blank){book.textStats.textShapes++;book.textStats.textCharacters+=text.length;book.textStats.paragraphs+=paragraphs;if(shape.ph)book.textStats.placeholders++;shape.textWritten=true;}
    return el('p:txBody',{},S.bodyPr(shape)+el('a:lstStyle')+parts.join(''));
  };
  T.styles=function(book,page,type,tag){let body='';for(let i=0;i<9;i++){const s=T.style(book,page,type,i);body+=T.pPr('a:lvl'+(i+1)+'pPr',s.p,s.r,book,page);}return el(tag,{},body);};
  T.media=async function(book,pkg){
    const media=new Map(),relationships=new Map();let number=0;
    for(const [index,rec]of book.textBlips)try{const image=await B.officeart.blip(B.officeart.record(rec.r,2)),name='image'+(++number)+'.'+image.ext;pkg.add('ppt/media/'+name,image.bytes,image.mime);media.set(index,name);}catch(e){if(e.name!=='LegacyError')throw e;}
    book.textStats.pictureBulletImages=media.size;
    book.bulletRel=index=>{const name=media.get(index);if(!name||!book.textPart)return null;const key=book.textPart+':'+index;if(!relationships.has(key))relationships.set(key,pkg.rel(book.textPart,'image','../media/'+name));return relationships.get(key);};
  };
  // A ChildAnchor uses its group's arbitrary coordinate system. It is not a
  // page coordinate in master units (modern writers often already use EMUs).
  const rectFor=shape=>{if(!shape.anchor)return null;const scale=shape.anchor.kind==='child'?v=>v:emu;return {x:scale(shape.anchor.left),y:scale(shape.anchor.top),cx:scale(shape.anchor.right-shape.anchor.left),cy:scale(shape.anchor.bottom-shape.anchor.top)};};
  T.tree=function(book,page,{layout,shapes,usedIds=[],fragment=false}={}){
    if(!page)return P.emptyShapeTree();
    const source=shapes||(layout?layout.textPlaceholders:page.textShapes||[]),used=new Set([1,...usedIds]);
    const ids=new Set();const collect=shapes=>{for(const s of shapes){if(s.id>1)ids.add(s.id);collect(s.children||[]);}};collect(source);let next=2;
    const identity=s=>{if(s.id>1&&!used.has(s.id)){used.add(s.id);return s.id;}while(ids.has(next)||used.has(next))next++;used.add(next);return next++;};
    const write=(shape,depth=0)=>{
      if(depth>B.limits.depth)B.fail('PPT_TEXT_LIMIT',N('Presentation text nesting exceeds the limit'));
      let rect=rectFor(shape);
      if(!rect&&shape.ph){let parent=page.kind==='notes'?book.notesMaster:page.parent;const seen=new Set();while(parent&&!seen.has(parent)){seen.add(parent);const match=parent.textAll?.find(s=>s.ph?.type===shape.ph.type&&s.anchor);if(match){rect=rectFor(match);break;}parent=parent.parent;}}
      const children=(shape.children||[]).map(s=>write(s,depth+1)).join('');
      if(children){if(!rect&&shape.coord)rect={x:emu(shape.coord.left),y:emu(shape.coord.top),cx:emu(shape.coord.right-shape.coord.left),cy:emu(shape.coord.bottom-shape.coord.top)};if(!rect){page.textReport('drawing','Group '+shape.id);return children;}
        const coord=shape.coord||null;
        return el('p:grpSp',{},el('p:nvGrpSpPr',{},el('p:cNvPr',{id:identity(shape),name:'Group '+shape.id})+el('p:cNvGrpSpPr')+el('p:nvPr'))+el('p:grpSpPr',{},S.transform({...shape,coord},rect,true))+children);}
      if(!shape.text&&!shape.ph)return '';if(!rect){if(shape.text?.text)page.textReport('text','Text box '+shape.id+' has no anchor');return '';}
      const id=identity(shape),ph=shape.ph?el('p:ph',{type:shape.ph.type,idx:shape.ph.idx,sz:shape.ph.sz,orient:shape.ph.orient}):'';
      let properties=S.transform(shape,rect)+el('a:prstGeom',{prst:'rect'},el('a:avLst'))+el('a:noFill')+el('a:ln',{},el('a:noFill'));
      // Text boxes and text placeholders retain their frame. Other drawing
      // objects get their text here and their complete geometry in D3.
      if(shape.type===202||shape.ph)try{properties=S.transform(shape,rect)+S.geometry(shape)+S.style(shape,{scheme:page.scheme});shape.frameWritten=true;}catch(e){if(e.name!=='LegacyError')throw e;page.textReport('drawing',e);}
      const nv=el('p:nvSpPr',{},el('p:cNvPr',{id,name:(shape.ph?'Placeholder ':'TextBox ')+id})+el('p:cNvSpPr',{txBox:shape.type===202?1:null},shape.ph?el('a:spLocks',{noGrp:1}):'')+el('p:nvPr',{},ph));
      return el('p:sp',{},nv+el('p:spPr',{},properties)+T.body(shape,book,page,!!layout));
    };
    const content=source.map(s=>write(s)).join('');return fragment?content:P.emptyShapeTree().replace('</p:spTree>',content+'</p:spTree>');
  };
  T.storedLayout=function(book,layout){
    const page=layout.source,base=layout.master.source,xml=layout.stored.xml;
    if(!page||page===base||page.compositeMasterId==null)return xml;
    const key=s=>JSON.stringify([s.text?.text,s.anchor?.left,s.anchor?.top,s.anchor?.right,s.anchor?.bottom]);
    const inherited=new Set((base?.textAll||[]).filter(s=>!s.ph).map(key));
    let candidates=page.textAll.filter(s=>!s.ph&&s.text?.text&&!inherited.has(key(s)));if(!candidates.length)return xml;
    const ids=[],existing=[],stack=[];let insert,shapeTree;
    const local=n=>n.name.split(':').at(-1),namespace=n=>n.ns[n.name.includes(':')?n.name.split(':')[0]:''];
    const visit=(node,depth)=>{const name=local(node),ns=namespace(node);if(ns===B.ns.p&&name==='spTree')shapeTree=node;if(ns===B.ns.p&&name==='cNvPr')ids.push(+node.attrs.id);if(ns===B.ns.p&&name==='sp')stack.push({node,text:'',ph:false,depth});if(ns===B.ns.p&&name==='ph'&&stack.length)stack.at(-1).ph=true;};
    visit.text=(text,node)=>{if(node&&namespace(node)===B.ns.a&&local(node)==='t'&&stack.length)stack.at(-1).text+=text;};
    visit.end=(node,close)=>{if(stack.at(-1)?.node===node){const s=stack.pop();if(!s.ph)existing.push(s.text);}if(node===shapeTree)insert=close;};
    B.checkXml(xml,visit);
    // A composite master can contain layout-owned static text absent from its
    // original master. Keep it once; never render layout placeholder prompts.
    const counts=new Map();for(const s of existing)counts.set(s,(counts.get(s)||0)+1);
    candidates=candidates.filter(s=>{const n=counts.get(s.text.text)||0;if(n){counts.set(s.text.text,n-1);return false;}return true;});
    if(!candidates.length||insert==null)return xml;
    const keep=new Set(candidates),filter=shapes=>shapes.flatMap(s=>{const children=filter(s.children||[]);return keep.has(s)||children.length?[{...s,children}]:[];});
    const added=T.tree(book,page,{shapes:filter(page.textShapes),usedIds:ids,fragment:true});return xml.slice(0,insert)+added+xml.slice(insert);
  };
})(typeof window!=='undefined'?window:globalThis);
