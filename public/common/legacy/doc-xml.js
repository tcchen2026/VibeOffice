/* WordprocessingML formatting, using the binary properties decoded by doc-sprm.
 * MS-DOC property definitions and ECMA-376 Part 1 §17.3/17.6/17.7/17.9.
 * XML ordering and defaults follow Office-authored WordprocessingML files.
 * Derived from Open XML SDK data/schemas/schemas_openxmlformats_org_wordprocessingml_2006_main.json (MIT): property ordering. */
(function(root){
  'use strict';
  const B=root.L.legacy,D=B.doc,el=B.el,N=s=>s;
  D.w=(name,attrs={},body)=>el('w:'+name,Object.fromEntries(Object.entries(attrs).map(([k,v])=>[k.includes(':')?k:'w:'+k,v])),body);
  const w=D.w,v=(name,value)=>value==null?'':w(name,{val:value}),flag=(name,value)=>value==null?'':w(name,{val:value?1:0});
  D.part=(name,body)=>B.xmlDeclaration+el('w:'+name,{'xmlns:w':B.ns.w,'xmlns:r':B.ns.officeRel.slice(0,-1)},body);
  D.bound=function(book,value,low,high,where){if(value==null)return value;if(value>=low&&value<=high)return Math.round(value);D.notice(book,'legacy-word-format-value',N('An invalid formatting value was replaced'),where,true);return Math.max(low,Math.min(high,Math.round(value)||low));};
  const underlines={0:'none',1:'single',2:'words',3:'double',4:'dotted',6:'thick',7:'dash',9:'dotDash',10:'dotDotDash',11:'wave',20:'dottedHeavy',23:'dashedHeavy',25:'dashDotHeavy',26:'dashDotDotHeavy',27:'wavyHeavy',39:'dashLong',43:'wavyDouble',55:'dashLongHeavy'};
  const borderTypes=['nil','single','thick','double','hairline','dotted','dashed','dotDash','dotDotDash','triple','thinThickSmallGap','thickThinSmallGap','thinThickThinSmallGap','thinThickMediumGap','thickThinMediumGap','thinThickThinMediumGap','thinThickLargeGap','thickThinLargeGap','thinThickThinLargeGap','wave','doubleWave','dashSmallGap','dashDotStroked','threeDEmboss','threeDEngrave','outset','inset'];
  // MS-DOC Ipat: indices 26–34 are reserved. The extra percentage patterns
  // start at 35; only those explicitly mapped to ST_Shd survive in OOXML.
  const patterns=Object.assign(['clear','solid','pct5','pct10','pct20','pct25','pct30','pct40','pct50','pct60','pct70','pct75','pct80','pct90','horzStripe','vertStripe','reverseDiagStripe','diagStripe','horzCross','diagCross','thinHorzStripe','thinVertStripe','thinReverseDiagStripe','thinDiagStripe','thinHorzCross','thinDiagCross'],{37:'pct12',38:'pct15',43:'pct35',44:'pct37',46:'pct45',49:'pct55',51:'pct62',52:'pct65',57:'pct85',58:'pct87',60:'pct95',65535:'nil'});
  D.borderXML=function(book,side,b,where){
    const val=borderTypes[b.type]||(b.type===255?'nil':null);
    if(!val){D.notice(book,'legacy-word-border',N('A decorative border was not kept'),where,true);return '';}
    return w(side,{val,sz:val==='nil'?null:D.bound(book,b.size,2,96,where),space:b.space,color:b.color,shadow:b.shadow?1:null,frame:b.frame?1:null});
  };
  D.shadingXML=function(book,p,where){const val=patterns[p.pattern];if(!val)D.notice(book,'legacy-word-shading',N('A shading pattern was not kept; its background color was retained'),where,true);return w('shd',{val:val||'clear',color:p.color,fill:p.fill});};
  D.formatWarnings=function(book,p,where){
    if(p.unsupported?.length)D.notice(book,'legacy-word-format','Unsupported binary formatting properties',where,false,p.unsupported.map(n=>n.toString(16)).join(', '));
    if(p.revisionFormatting)D.notice(book,'legacy-word-revisions',N('Revision history was not kept'),where,true);
  };
  D.runXML=function(book,p,where,styleBase,directBase){
    // Direct run properties are absolute, unlike style toggles. Omit inherited
    // values so changing a retained style still changes its unmodified runs.
    if(directBase){p={...p};for(const key of Object.keys(p))if(!['charStyle','highlight','rtl','complexScript'].includes(key)&&JSON.stringify(p[key])===JSON.stringify(directBase[key]))delete p[key];}
    if(styleBase){p={...p};for(const key of Object.keys(p))if(typeof p[key]!=='boolean'&&JSON.stringify(p[key])===JSON.stringify(styleBase[key]))delete p[key];}
    // These binary style properties cannot occur in OOXML StyleRunProperties.
    // Keep them on the affected runs and paragraph marks instead.
    if(styleBase){delete p.highlight;delete p.charStyle;delete p.rtl;delete p.complexScript;}
    let out=book.styles[p.charStyle]?.kind===2?v('rStyle',book.styles[p.charStyle].id):'';const fonts={ascii:book.fonts[p.font]?.name,hAnsi:book.fonts[p.fontOther]?.name,eastAsia:book.fonts[p.fontEastAsia]?.name,cs:book.fonts[p.fontCs]?.name,hint:['default','eastAsia','cs'][p.hint]};
    if(Object.values(fonts).some(x=>x!=null))out+=w('rFonts',fonts);
    const toggles=[['bold','b'],['boldCs','bCs'],['italic','i'],['italicCs','iCs'],['caps','caps'],['smallCaps','smallCaps'],['strike','strike'],['doubleStrike','dstrike'],['outline','outline'],['shadow','shadow'],['emboss','emboss'],['imprint','imprint'],['noProof','noProof']];
    for(const [key,name] of toggles)out+=styleBase?(!!p[key]!==!!styleBase[key]?w(name):''):flag(name,p[key]);
    out+=flag('snapToGrid',p.snapToGrid);
    for(const [key,name] of [['vanish','vanish'],['webHidden','webHidden']])out+=styleBase?(!!p[key]!==!!styleBase[key]?w(name):''):flag(name,p[key]);
    out+=v('color',p.color)+v('spacing',p.spacing)+v('w',p.scale==null?null:D.bound(book,p.scale,1,600,where))+v('kern',p.kern)+v('position',p.position==null?null:D.bound(book,p.position,-31680,31680,where));
    out+=v('sz',p.size==null?null:D.bound(book,p.size,2,1638,where))+v('szCs',p.sizeCs==null?null:D.bound(book,p.sizeCs,2,1638,where));
    if(p.highlight!=null)out+=v('highlight',['none','black','blue','cyan','green','magenta','red','yellow','white','darkBlue','darkCyan','darkGreen','darkMagenta','darkRed','darkYellow','darkGray','lightGray'][p.highlight]||'none');
    if(p.underline!=null){const val=underlines[p.underline];if(val)out+=w('u',{val,color:p.underlineColor});else D.notice(book,'legacy-word-underline',N('An underline style was not kept'),where,true);}
    if(p.border)out+=D.borderXML(book,'bdr',p.border,where);if(p.shading)out+=D.shadingXML(book,p.shading,where);
    out+=v('vertAlign',['baseline','superscript','subscript'][p.vertAlign])+flag('rtl',p.rtl)+flag('cs',p.complexScript);
    const lang={val:B.languageIds?.[p.language],eastAsia:B.languageIds?.[p.languageEastAsia],bidi:B.languageIds?.[p.languageCs]};if(Object.values(lang).some(Boolean))out+=w('lang',lang);
    out+=flag('specVanish',p.specVanish);
    return out;
  };
  D.paragraphXML=function(book,p,where,options={}){
    if(options.style){const inherited=options.style.p;p={...p};for(const key of Object.keys(p))if(JSON.stringify(p[key])===JSON.stringify(inherited[key]))delete p[key];}
    let out='';if(options.style)out+=v('pStyle',options.style.id);
    out+=flag('keepNext',p.keepNext)+flag('keepLines',p.keepLines)+flag('pageBreakBefore',p.pageBreakBefore);
    if(p.frameWidth!=null||p.frameX!=null||p.frameY!=null){
      const vertical=(p.frameAnchor||0)&3,horizontal=((p.frameAnchor||0)>>>2)&3;
      out+=w('framePr',{w:p.frameWidth,h:p.frameHeight==null?null:p.frameHeight&32767,hRule:p.frameHeight==null?null:p.frameHeight&32768?'exact':'atLeast',hSpace:p.frameHSpace,vSpace:p.frameVSpace,wrap:['auto','notBeside','around','tight','through','none'][p.frameWrap],vAnchor:['margin','page','text'][vertical],hAnchor:['text','margin','page'][horizontal],
        x:p.frameX>=0?p.frameX:null,xAlign:p.frameX<0?({[-4]:'center',[-8]:'right',[-12]:'inside',[-16]:'outside',[-0]:'left'})[p.frameX]:null,
        y:p.frameY>=0?p.frameY:null,yAlign:p.frameY<0?({[-4]:'top',[-8]:'center',[-12]:'bottom',[-16]:'inside',[-20]:'outside'})[p.frameY]:null});
    }
    out+=flag('widowControl',p.widowControl);
    const list=options.noList?null:D.listFor(book,p,where);
    if(list)out+=w('numPr',{},v('ilvl',list.level)+v('numId',list.override.id));else if(p.list===0&&options.explicitList)out+=w('numPr',{},v('numId',0));
    out+=flag('suppressLineNumbers',p.suppressLineNumbers);
    if(p.borders)out+=w('pBdr',{},['top','left','bottom','right','between','bar'].filter(side=>p.borders[side]).map(side=>D.borderXML(book,side,p.borders[side],where)).join(''));
    if(p.shading)out+=D.shadingXML(book,p.shading,where);
    if(p.tabs?.length)out+=w('tabs',{},p.tabs.slice(0,64).map(t=>w('tab',{val:t.clear?'clear':['left','center','right','decimal','bar','num'][t.align]||'left',leader:t.clear?null:['none','dot','hyphen','underscore','heavy','middleDot'][t.leader],pos:D.bound(book,t.position,-31680,31680,where)})).join(''));
    for(const key of ['suppressAutoHyphens','kinsoku','wordWrap','overflowPunct','topLinePunct','autoSpaceDE','autoSpaceDN','bidi','adjustRightInd','snapToGrid'])out+=flag(key,p[key]);
    const spacing={before:p.before,after:p.after,beforeLines:p.beforeLines,afterLines:p.afterLines,beforeAutospacing:p.beforeAuto==null?null:+p.beforeAuto,afterAutospacing:p.afterAuto==null?null:+p.afterAuto};
    if(p.line){spacing.line=Math.abs(p.line.value);spacing.lineRule=p.line.multiple?'auto':p.line.value<0?'exact':'atLeast';}
    if(Object.values(spacing).some(x=>x!=null))out+=w('spacing',spacing);
    const ind={left:p.left,right:p.right,leftChars:p.leftChars,rightChars:p.rightChars,firstLine:p.first>=0?p.first:null,hanging:p.first<0?-p.first:null,firstLineChars:p.firstChars>=0?p.firstChars:null,hangingChars:p.firstChars<0?-p.firstChars:null};
    if(Object.values(ind).some(x=>x!=null))out+=w('ind',ind);
    out+=flag('contextualSpacing',p.contextualSpacing)+flag('mirrorIndents',p.mirrorIndents);
    let align=p.alignLogical??p.align;if(p.alignLogical!=null&&p.bidi&&align<3)align=2-align;
    out+=v('jc',['left','center','right','both','distribute','mediumKashida','numTab','highKashida','lowKashida','thaiDistribute'][align]);
    out+=v('textDirection',['lrTb','tbRl','btLr','lrTbV','tbRlV'][p.textDirection])+v('textAlignment',['top','center','baseline','bottom','auto'][p.textAlignment])+v('outlineLvl',p.outlineLevel==null?null:Math.min(9,p.outlineLevel));
    return out;
  };
  const hex=bytes=>Array.from(bytes,b=>b.toString(16).padStart(2,'0')).join('').toUpperCase();
  D.fontsXML=function(book){
    const used=new Set();return D.part('fonts',book.fonts.map(f=>{
      if(!f.name||used.has(f.name.toLowerCase()))return '';used.add(f.name.toLowerCase());
      const sig=new B.Reader(f.signature);return w('font',{name:f.name},v('altName',f.alternate)+v('panose1',hex(f.panose))+v('charset',f.charset.toString(16).padStart(2,'0').toUpperCase())+v('family',['auto','roman','swiss','modern','script','decorative'][f.family]||'auto')+v('pitch',['default','fixed','variable'][f.pitch]||'default')+w('sig',Object.fromEntries(['usb0','usb1','usb2','usb3','csb0','csb1'].map((key,i)=>[key,sig.u32(i*4).toString(16).padStart(8,'0').toUpperCase()]))));
    }).join(''));
  };
  D.stylesXML=function(book){
    let body=w('docDefaults',{},w('rPrDefault',{},w('rPr',{},D.runXML(book,book.defaults,'Default text')))+w('pPrDefault'));
    for(const s of book.styles){
      if(!s||s.kind!==1&&s.kind!==2)continue;
      const parent=book.styles[s.basedOn],next=book.styles[s.next],link=book.styles[s.link];
      let content=v('name',s.name)+(parent&&parent.kind===s.kind&&parent!==s?v('basedOn',parent.id):'')+(next&&next.kind===s.kind&&s.kind===1?v('next',next.id):'')+(link&&link.kind!==s.kind&&(link.kind===1||link.kind===2)?v('link',link.id):'');
      if(s.flags&1)content+=w('autoRedefine');if(s.flags&2)content+=w('hidden');if(s.flags&0x100)content+=w('semiHidden');if(s.flags&0x800)content+=w('unhideWhenUsed');if(s.flags&0x1000)content+=w('qFormat');if(s.flags&0x200)content+=w('locked');
      if(s.kind===1){const p=D.paragraphXML(book,s.p,'Style '+s.name);if(p)content+=w('pPr',{},p);}
      const r=D.runXML(book,s.c,'Style '+s.name,parent?.c||book.defaults);if(r)content+=w('rPr',{},r);
      body+=w('style',{type:s.kind===1?'paragraph':'character',default:s.sti===0||s.sti===65?1:null,customStyle:s.sti===4094?1:null,styleId:s.id},content);
    }
    return D.part('styles',body);
  };
  D.numberingXML=function(book){
    function level(l,index,list){
      const where='List '+(list.id+1),format=D.numberFormats[l.format];
      if(!format||format==='application')D.notice(book,'legacy-word-number-format',N('A numbering style was replaced with decimal numbers'),where,true);
      let text=l.text;
      if(l.format!==23){const offsets=new Set(l.positions.filter(Boolean).map(n=>n-1));text=Array.from({length:text.length},(_,i)=>{const ch=text.charCodeAt(i);return offsets.has(i)&&ch<9?(list.levels[ch]?.format===255?'':'%'+(ch+1)):text[i];}).join('');}
      const style=book.styles[list.styles[index]],body=v('start',D.bound(book,l.start,0,2147483647,where))+v('numFmt',format&&format!=='application'?format:'decimal')+((l.flags&8)?v('lvlRestart',0):l.restart&&l.restart<=index?v('lvlRestart',l.restart):'')+(style?.kind===1?v('pStyle',style.id):'')+((l.flags&4)?w('isLgl'):'')+v('suff',['tab','space','nothing'][l.follow]||'tab')+v('lvlText',text)+v('lvlJc',['left','center','right'][l.flags&3]||'left')+w('pPr',{},D.paragraphXML(book,l.p,where,{noList:true}))+w('rPr',{},D.runXML(book,l.c,where));
      return w('lvl',{ilvl:index,tentative:l.flags&128?1:null},body);
    }
    let body='';for(const l of book.lists)body+=w('abstractNum',{abstractNumId:l.id},v('nsid',l.lsid.toString(16).padStart(8,'0').toUpperCase())+v('multiLevelType',l.flags&1?'singleLevel':l.flags&16?'hybridMultilevel':'multilevel')+v('tmpl',l.template.toString(16).padStart(8,'0').toUpperCase())+l.levels.map((x,i)=>level(x,i,l)).join(''));
    for(const o of book.overrides)if(o.list)body+=w('num',{numId:o.id},v('abstractNumId',o.list.id)+o.levels.map(l=>w('lvlOverride',{ilvl:l.index},v('startOverride',l.start==null?null:D.bound(book,l.start,0,2147483647,'List '+o.id))+(l.level?level(l.level,l.index,o.list):''))).join(''));
    return D.part('numbering',body);
  };
})(typeof window!=='undefined'?window:globalThis);
