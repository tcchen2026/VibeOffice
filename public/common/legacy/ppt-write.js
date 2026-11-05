/* Binary PowerPoint -> PresentationML structure. Text/drawing content is added
 * by the subsequent converter stages; until then every detected omission is
 * reported. ECMA-376 packaging and Office-authored master/layout conventions. */
(function(root){
  'use strict';
  const B=root.L.legacy,P=B.ppt,el=B.el,N=s=>s;
  const ct='application/vnd.openxmlformats-officedocument.',types={pptx:ct+'presentationml.presentation.main+xml',ppsx:ct+'presentationml.slideshow.main+xml',potx:ct+'presentationml.template.main+xml'};
  const ns=()=>({'xmlns:a':B.ns.a,'xmlns:r':B.ns.officeRel.slice(0,-1),'xmlns:p':B.ns.p});
  const part=(tag,body,attrs={})=>B.xmlDeclaration+el('p:'+tag,{...ns(),...attrs},body);
  P.emptyShapeTree=()=>el('p:spTree',{},el('p:nvGrpSpPr',{},el('p:cNvPr',{id:1,name:''})+el('p:cNvGrpSpPr')+el('p:nvPr'))+el('p:grpSpPr',{},el('a:xfrm',{},el('a:off',{x:0,y:0})+el('a:ext',{cx:0,cy:0})+el('a:chOff',{x:0,y:0})+el('a:chExt',{cx:0,cy:0}))));
  const fallback=()=>({scheme:P.defaultScheme,backgroundScheme:P.defaultScheme,colorMap:P.colorMap,children:[],schemeGallery:[],flags:0});
  const background=()=>el('p:bg',{},el('p:bgPr',{},el('a:solidFill',{},el('a:schemeClr',{val:'bg1'}))+el('a:effectLst')));
  const commonSlide=(page,bg,book,layout)=>el('p:cSld',{name:page.name||null},(bg?background():'')+P.text.tree(book,page,{layout}));
  const override=page=>el('p:clrMapOvr',{},page?.hasColorOverride?el('a:overrideClrMapping',page.colorMap):el('a:masterClrMapping'));
  P.write=async function(book,cfb,options={}){
    const losses=book.losses||[],pkg=new B.Package(),pres='ppt/presentation.xml';let themeCount=0,overrideCount=0;
    await P.prepareDesigns(book,losses);
    P.text.prepare(book);
    await P.text.embedded(book);
    await P.text.media(book,pkg);
    const theme=(from,page)=>{const stored=page.theme?.root==='theme'?page.theme.xml:null,xml=stored||P.themeXml(page,book),name='theme'+(++themeCount)+'.xml';pkg.add('ppt/theme/'+name,xml,ct+'theme+xml');pkg.rel(from,'theme','../theme/'+name);};
    const themeOverride=(from,page,parent)=>{
      const supplied=page.theme?.root==='themeOverride'?page.theme.xml:null;
      const different=parent&&page.scheme.join()!==parent.scheme.join();
      page.hasColorOverride=page.hasStoredColorMap||different;
      if(!supplied&&!different)return;
      const xml=supplied||B.xmlDeclaration+el('a:themeOverride',{'xmlns:a':B.ns.a},P.schemeXml(page.scheme));
      const name='themeOverride'+(++overrideCount)+'.xml';pkg.add('ppt/theme/'+name,xml,ct+'themeOverride+xml');pkg.rel(from,'themeOverride','../theme/'+name);
    };
    const masterIds=[];
    for(const m of book.outputMasters){
      const page=m.source||fallback(),name='ppt/slideMasters/slideMaster'+m.index+'.xml',ids=[];
      book.textPart=name;
      for(const layout of m.layouts){const rid=pkg.rel(name,'slideLayout','../slideLayouts/slideLayout'+layout.index+'.xml');ids.push(el('p:sldLayoutId',{id:layout.id,'r:id':rid}));}
      theme(name,page);pkg.add(name,part('sldMaster',commonSlide(page,true,book)+el('p:clrMap',page.colorMap)+el('p:sldLayoutIdLst',{},ids.join(''))+(page.storedTextStyles||el('p:txStyles',{},P.text.styles(book,m.source,0,'p:titleStyle')+P.text.styles(book,m.source,1,'p:bodyStyle')+P.text.styles(book,m.source,4,'p:otherStyle')))),ct+'presentationml.slideMaster+xml');
      masterIds.push(el('p:sldMasterId',{id:m.id,'r:id':pkg.rel(pres,'slideMaster','slideMasters/slideMaster'+m.index+'.xml')}));
    }
    for(const layout of book.outputLayouts){
      const name='ppt/slideLayouts/slideLayout'+layout.index+'.xml',page=layout.source||fallback(),parent=layout.master.source||fallback();
      book.textPart=name;
      pkg.rel(name,'slideMaster','../slideMasters/slideMaster'+layout.master.index+'.xml');
      if(!layout.stored)themeOverride(name,page,parent);
      const titleMaster=page.kind==='master'&&page.record?.type===1006;
      pkg.add(name,layout.stored?P.text.storedLayout(book,layout):part('sldLayout',commonSlide(page,titleMaster&&!(page.flags&4),book,titleMaster?null:layout)+override(page),{type:layout.type,preserve:1,showMasterSp:titleMaster&&!(page.flags&1)?0:null}),ct+'presentationml.slideLayout+xml');
    }
    let notesId='',handoutId='';
    if(book.notesMaster||book.slides.some(s=>s.notes)){
      const page=book.notesMaster||fallback(),name='ppt/notesMasters/notesMaster1.xml';theme(name,page);
      book.textPart=name;
      pkg.add(name,part('notesMaster',commonSlide(page,true,book)+el('p:clrMap',page.colorMap)+(page.storedNotesStyle||P.text.styles(book,book.notesMaster,2,'p:notesStyle'))),ct+'presentationml.notesMaster+xml');
      notesId=el('p:notesMasterIdLst',{},el('p:notesMasterId',{'r:id':pkg.rel(pres,'notesMaster','notesMasters/notesMaster1.xml')}));
    }
    if(book.handoutMaster){
      const page=book.handoutMaster,name='ppt/handoutMasters/handoutMaster1.xml';theme(name,page);
      book.textPart=name;
      pkg.add(name,part('handoutMaster',commonSlide(page,true,book)+el('p:clrMap',page.colorMap)),ct+'presentationml.handoutMaster+xml');
      handoutId=el('p:handoutMasterIdLst',{},el('p:handoutMasterId',{'r:id':pkg.rel(pres,'handoutMaster','handoutMasters/handoutMaster1.xml')}));
    }
    const sourceIds=new Set(book.slides.map(s=>s.id).filter(id=>id>=256&&id<=0x7fffffff)),usedIds=new Set(),slideIds=[];let nextId=256,noteCount=0;
    for(let i=0;i<book.slides.length;i++){
      const slide=book.slides[i],name='ppt/slides/slide'+(i+1)+'.xml',layout=slide.outputLayout;let id=slide.id;
      book.textPart=name;
      if(id<256||id>0x7fffffff||usedIds.has(id)){while(sourceIds.has(nextId)||usedIds.has(nextId))nextId++;id=nextId++;}usedIds.add(id);
      pkg.rel(name,'slideLayout','../slideLayouts/slideLayout'+layout.index+'.xml');
      themeOverride(name,slide,layout.source||fallback());
      pkg.add(name,part('sld',commonSlide(slide,!(slide.flags&4),book)+override(slide),{show:slide.hidden?0:null,showMasterSp:slide.flags&1?null:0}),ct+'presentationml.slide+xml');
      slideIds.push(el('p:sldId',{id,'r:id':pkg.rel(pres,'slide','slides/slide'+(i+1)+'.xml')}));
      if(slide.notes){
        const note=slide.notes,base='notesSlide'+(++noteCount)+'.xml',file='ppt/notesSlides/'+base;
        book.textPart=file;
        pkg.rel(name,'notesSlide','../notesSlides/'+base);pkg.rel(file,'slide','../slides/slide'+(i+1)+'.xml');pkg.rel(file,'notesMaster','../notesMasters/notesMaster1.xml');
        themeOverride(file,note,book.notesMaster||fallback());
        pkg.add(file,part('notes',commonSlide(note,!(note.flags&4),book)+override(note),{showMasterSp:note.flags&1?null:0}),ct+'presentationml.notesSlide+xml');
      }
    }
    const sizeType=['screen4x3','letter','A4','35mm','overhead','banner',null][book.sizeType];
    const extension=String(options.extension||'ppt').replace(/^\./,'').toLowerCase(),type=extension==='pps'?'ppsx':extension==='pot'?'potx':'pptx';
    pkg.add(pres,part('presentation',el('p:sldMasterIdLst',{},masterIds.join(''))+notesId+handoutId+(slideIds.length?el('p:sldIdLst',{},slideIds.join('')):'')+el('p:sldSz',{...book.size,type:sizeType})+el('p:notesSz',book.notesSize),{saveSubsetFonts:1,firstSlideNum:book.firstSlideNumber!==1&&book.firstSlideNumber<10000?book.firstSlideNumber:null,rtl:book.rtl?1:null}),types[type]);
    pkg.rel('','officeDocument',pres);
    losses.push(...B.addProperties(pkg,cfb,{extended:{Slides:book.slides.length,Notes:noteCount,HiddenSlides:book.slides.filter(s=>s.hidden).length}}));
    // Scope is deliberately measurable: D1 does not claim content fidelity.
    // Later stages replace these notices as they start writing each feature.
    let textCharacters=0;
    const backgroundSources=new Set([...book.outputMasters.map(m=>m.source),...book.outputLayouts.filter(l=>!l.stored).map(l=>l.source)]);
    for(const page of [...book.slides,...book.masters,...book.notes,book.notesMaster,book.handoutMaster].filter(Boolean)){
      const c=page.content; textCharacters+=c.text.reduce((n,t)=>n+t.length,0);
      const omittedShapes=Math.max(0,c.shapes-(page.textAll||[]).filter(s=>s.textWritten&&s.frameWritten).length);
      if(omittedShapes)P.notice(losses,'legacy-slide-drawings',N('Drawing objects were not kept'),page.where,true,omittedShapes+' objects; '+c.pictures+' pictures; '+c.tables+' tables');
      // A picture/gradient background is content even when the slide has no
      // ordinary shapes. Don't hide its omission behind an empty-shape count.
      const backgroundWritten=page.kind==='master'?backgroundSources.has(page):!(page.flags&4);
      if(page.backgroundShape&&backgroundWritten){const shape=page.backgroundShape,get=(id,fallback)=>shape.properties.get(id)?.value??fallback;let differs=get(384,0)!==0||get(386,65536)!==65536;try{differs ||= B.shapes.color(get(385,0x08000000),{scheme:page.scheme},shape)!==page.scheme[0];}catch(e){if(e.name!=='LegacyError')throw e;differs=true;}if(differs)P.notice(losses,'legacy-slide-background',N('The slide background was replaced with its color-scheme background'),page.where,true);}
      if(c.animations)P.notice(losses,'legacy-slide-animation',N('Animations were not kept'),page.where,true);
      if(c.links)P.notice(losses,'legacy-slide-links',N('Hyperlinks were not kept'),page.where,true);
      if(page.kind==='slide'||page.kind==='notes'){const missing=P.text.missingHeaders(book,page);if(missing.length)P.notice(losses,'legacy-slide-headers-footers',N('Some headers and footers were not kept'),page.where,true,missing.join(', '));}
      if(page.transition)P.notice(losses,'legacy-slide-transition',N('Slide transitions and automatic advance settings were not kept'),page.where,true);
      if(page.kind==='notes'&&!page.slide)P.notice(losses,'legacy-unattached-notes',N('Notes without a matching slide were not kept'),page.where,true);
    }
    if(book.embeddedFonts)P.notice(losses,'legacy-embedded-fonts',N('Embedded fonts were not kept'),'Presentation',true);
    const vba=book.documentRecords.find(r=>r.type===1023);
    if(vba&&P.children(vba,()=>P.notice(losses,'legacy-macros',N('Damaged macros were not kept'),'Presentation',true)).some(r=>r.type===1024&&r.r.length>=8&&r.r.u32(4)))P.notice(losses,'legacy-macros',N('Macros were not kept'),'Presentation',true);
    return {pkg:await pkg.finish(),type,losses,stats:{slides:book.slides.length,masters:book.outputMasters.length,sourceMasters:book.masters.length,layouts:book.outputLayouts.length,storedLayouts:book.outputLayouts.filter(l=>l.stored).length,notes:noteCount,handoutMasters:book.handoutMaster?1:0,fonts:book.fonts.size,themes:themeCount,themeOverrides:overrideCount,incrementalEdits:book.edits.length,sourceTextCharacters:textCharacters,...book.textStats}};
  };
  P.convert=async function(bytes,options={}){
    const cfb=B.cfb(bytes),user=cfb.stream('Current User');
    if(cfb.entries.some(e=>e.name.toLowerCase()==='pp40')||user&&user.length<28)B.fail('PPT_VERSION',N('PowerPoint 95 presentations cannot be opened yet'));
    const book=P.structure(cfb,[]);
    if(book.encrypted)B.fail(options.password?'ENCRYPTION_UNSUPPORTED':'PASSWORD_REQUIRED',options.password?N('Encrypted binary presentations cannot be opened yet'):N('This presentation needs a password'));
    return P.write(book,cfb,options);
  };
})(typeof window!=='undefined'?window:globalThis);
