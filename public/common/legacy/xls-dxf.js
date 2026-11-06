/* Derived from Apache POI hssf/record/{CFRuleBase,cf/FontFormatting,
 * cf/BorderFormatting,cf/PatternFormatting,common/ExtendedColor}.java (Apache-2.0).
 * MS-XLS DXFN/DXFN12/XFExtNoFRT add number formats, alignment and extended colors. */
(function(root){
  'use strict';
  const B=root.L.legacy,X=B.xls,el=B.el,bad=s=>B.fail('CF_FORMAT',s);
  // An omitted Theme payload means Excel's Office 2007 default theme (MS-XLS
  // Theme). Resolve its colors directly when there is no theme part to refer to.
  const defaultTheme=['FFFFFF','000000','EEECE1','1F497D','4F81BD','C0504D','9BBB59','8064A2','4BACC6','F79646','0000FF','800080'];
  const resolveTheme=(c,book)=>c?.theme!=null&&!book.theme?{rgb:'FF'+defaultTheme[c.theme],...(c.tint?{tint:c.tint}:{})}:c;
  X.cfColor=function(r,p,book){
    const color=resolveTheme(X.colorValue(r.u32(p),r.u32(p+4),true),book),tint=r.f64(p+8);
    if(!color||!Number.isFinite(tint)||Math.abs(tint)>1)bad('Invalid conditional-format color');if(tint)color.tint=tint;return color;
  };
  X.dxf=function(r,p,book,styles,end){
    const flags=r.u32(p),extra=r.u16(p+4);p+=6;let number='',font='',fill='',border='',alignment='',protection='';
    const fontItems={},sides={},fillAttrs={},fillColors={},borderAttrs={};
    if(flags&0x2000000){
      if(extra&1){const size=r.u16(p),s=X.string(r,p+2);if(size<2||s.end>p+size)bad('Invalid differential number format');number=el('numFmt',{numFmtId:164,formatCode:s.value});p+=size;}
      else{const id=r.u8(p+1);r.check(p,2);p+=2;if(!(flags&0x80000))number=el('numFmt',{numFmtId:id,formatCode:styles.formats.get(id)||builtinNumber(id)});}
    }
    if(flags&0x4000000){
      const f=r.reader(p,118);p+=118;const opts=f.u32(68),ignore=f.u32(88),size=f.i32(64),color=f.i32(80);
      if(f.u8(0))fontItems.name=el('name',{val:X.string(f,0,true,1200,true).value});
      if(size>0&&size<=8180)fontItems.sz=el('sz',{val:size/20});
      if(!(ignore&2)){fontItems.i=el('i',{val:opts&2?1:0});if(f.u16(72)<0x7fff)fontItems.b=el('b',{val:f.u16(72)>=700?1:0});}
      for(const [key,bit]of [['outline',8],['shadow',16],['strike',128]])if(!(ignore&bit))fontItems[key]=el(key,{val:opts&bit?1:0});
      if(f.u32(92)===0)fontItems.vertAlign=el('vertAlign',{val:['baseline','superscript','subscript'][f.u16(74)]||'baseline'});
      if(f.u32(96)===0)fontItems.u=el('u',{val:({0:'none',1:'single',2:'double',33:'singleAccounting',34:'doubleAccounting'})[f.u8(76)]||'none'});
      if(color>=0)fontItems.color=X.styleColor('color',color);
    }
    if(flags&0x8000000){
      const a=r.u32(p),relative=r.i32(p+4);p+=8;const attrs={};
      if(!(flags&1))attrs.horizontal=['general','left','center','right','fill','justify','centerContinuous','distributed'][a&7];
      if(!(flags&2))attrs.vertical=['top','center','bottom','justify','distributed'][a>>>4&7];
      if(!(flags&4))attrs.wrapText=a&8?1:0;if(!(flags&8)){const n=a>>>8&255;if(n<=180||n===255)attrs.textRotation=n;}
      if(!(flags&16))attrs.justifyLastLine=a&128?1:0;if(!(flags&32)){if(relative===255)attrs.indent=a>>>16&15;else if(Math.abs(relative)<=250)attrs.relativeIndent=relative;}
      if(!(flags&64))attrs.shrinkToFit=a&0x100000?1:0;if(!(flags&0x80000000)&&extra&0x8000)attrs.readingOrder=Math.min(2,a>>>22&3);alignment=el('alignment',attrs);
    }
    if(flags&0x10000000){
      const a=r.u32(p),b=r.u32(p+4);p+=8;
      const values={left:[a&15,a>>>16&127],right:[a>>>4&15,a>>>23&127],top:[a>>>8&15,b&127],bottom:[a>>>12&15,b>>>7&127],diagonal:[b>>>21&15,b>>>14&127]};
      ['left','right','top','bottom'].forEach((tag,i)=>{if(!(flags&(0x400<<i)))sides[tag]=values[tag];});
      if(!(flags&0x4000))borderAttrs.diagonalDown=a&0x40000000?1:0;if(!(flags&0x8000))borderAttrs.diagonalUp=a&0x80000000?1:0;
      if((flags&0xc000)!==0xc000)sides.diagonal=values.diagonal;
      if(extra&4)borderAttrs.outline=1;
    }
    if(flags&0x20000000){
      const a=r.u16(p),b=r.u16(p+2);p+=4;
      if(!(flags&0x10000))fillAttrs.patternType=X.stylePatterns[a>>>10&63]||'none';
      if(!(flags&0x20000))fillColors.fgColor=b&127;if(!(flags&0x40000))fillColors.bgColor=b>>>7&127;
    }
    if(flags&0x40000000){const n=r.u16(p);p+=2;protection=el('protection',{locked:flags&256?null:n&1?1:0,hidden:flags&512?null:n&2?1:0});}
    if(end!=null&&p<end){
      // XFExtNoFRT: reserved, source XF index (ignored for DXF), reserved, count.
      const count=r.u16(p+6);p+=8;
      for(let i=0;i<count;i++){
        const id=r.u16(p),size=r.u16(p+2);if(size<4||p+size>end)bad('Invalid differential format extension');const d=r.reader(p+4,size-4);p+=size;
        if([4,5,7,8,9,10,11,13].includes(id)){
          const color=resolveTheme(X.extColor(d,0,true),book);if(!color)bad('Unknown differential color');
          if(id===13)fontItems.color=el('color',color);else if(id===4||id===5)fillColors[id===4?'fgColor':'bgColor']=color;
          else{const tag=({7:'top',8:'bottom',9:'left',10:'right',11:'diagonal'})[id];if(sides[tag])sides[tag][1]=color;}
        }else if(id===14){const val=d.length===1?d.u8(0):d.u16(0);if(book.theme&&val<=2)fontItems.scheme=el('scheme',{val:['none','major','minor'][val]});}
        else if(id===15){/* No live CF fixture uses extended indentation. */bad('Extended differential indentation is not supported');}
        else bad('Unsupported differential format extension');
      }
      if(p!==end)bad('Invalid differential format size');
    }
    font=Object.keys(fontItems).length?el('font',{},X.fontOrder.map(key=>fontItems[key]||'').join('')):'';
    if(flags&0x20000000||Object.keys(fillColors).length)fill=el('fill',{},el('patternFill',fillAttrs,Object.entries(fillColors).map(([tag,c])=>X.styleColor(tag,c)).join('')));
    if(flags&0x10000000)border=el('border',borderAttrs,Object.entries(sides).map(([tag,[line,c]])=>el(tag,{style:X.borderLines[line]||null},line?X.styleColor('color',c):undefined)).join(''));
    return {body:font+number+fill+alignment+border+protection,end:p};
  };
  // Built-in formats stored by index in DXFNumIFmt. Non-built-ins use DXFNumUsr.
  const builtinNumber=id=>({0:'General',1:'0',2:'0.00',3:'#,##0',4:'#,##0.00',9:'0%',10:'0.00%',11:'0.00E+00',12:'# ?/?',13:'# ??/??',14:'mm-dd-yy',15:'d-mmm-yy',16:'d-mmm',17:'mmm-yy',18:'h:mm AM/PM',19:'h:mm:ss AM/PM',20:'h:mm',21:'h:mm:ss',22:'m/d/yy h:mm',37:'#,##0 ;(#,##0)',38:'#,##0 ;[Red](#,##0)',39:'#,##0.00;(#,##0.00)',40:'#,##0.00;[Red](#,##0.00)',45:'mm:ss',46:'[h]:mm:ss',47:'mmss.0',48:'##0.0E+0',49:'@'})[id]??bad('Unsupported built-in differential number format');
})(typeof window!=='undefined'?window:globalThis);
