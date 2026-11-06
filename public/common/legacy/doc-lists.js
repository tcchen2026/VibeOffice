/* Word list definitions and overrides (MS-DOC PlfLst, LVL, PlfLfo).
 * Derived from Apache POI hwpf/model/{ListTables,ListLevel,PlfLfo,LFOData,
 * ListFormatOverrideLevel}.java (Apache-2.0).
 * Derived from Apache OpenOffice sw/source/filter/ww8/ww8par3.cxx (Apache-2.0). */
(function(root){
  'use strict';
  const B=root.L.legacy,D=B.doc,N=s=>s;
  D.numberFormats=['decimal','upperRoman','lowerRoman','upperLetter','lowerLetter','ordinal','cardinalText','ordinalText','hex','chicago','ideographDigital','japaneseCounting','aiueo','iroha','decimalFullWidth','decimalHalfWidth','japaneseLegal','japaneseDigitalTenThousand','decimalEnclosedCircle','decimalFullWidth2','aiueoFullWidth','irohaFullWidth','decimalZero','bullet','ganada','chosung','decimalEnclosedFullstop','decimalEnclosedParen','decimalEnclosedCircleChinese','ideographEnclosedCircle','ideographTraditional','ideographZodiac','ideographZodiacTraditional','taiwaneseCounting','ideographLegalTraditional','taiwaneseCountingThousand','taiwaneseDigital','chineseCounting','chineseLegalSimplified','chineseCountingThousand','application','koreanDigital','koreanCounting','koreanLegal','koreanDigital2','hebrew1','arabicAlpha','hebrew2','arabicAbjad','hindiVowels','hindiConsonants','hindiNumbers','hindiCounting','thaiLetters','thaiNumbers','thaiCounting','vietnameseCounting','numberInDash','russianLower','russianUpper'];
  D.numberFormats[255]='none';
  function level(book,r,at){
    const f=r.reader(at,28,'LVLF'),np=f.u8(25),nc=f.u8(24);let p=at+28;
    const pOps=D.sprms(r.reader(p,np,'List paragraph properties'));p+=np;
    const cOps=D.sprms(r.reader(p,nc,'List character properties'));p+=nc;
    const chars=r.u16(p);p+=2;const text=r.text(p,chars*2,1200);p+=chars*2;
    return {end:p,start:f.i32(0),format:f.u8(4),flags:f.u8(5),positions:[...f.sub(6,9)],follow:f.u8(15),restart:f.u8(26),text,p:D.apply(pOps,1),c:D.apply(cOps,2),cOps};
  }
  D.lists=function(book){
    book.lists=[];book.overrides=[];
    D.optional(book,'Lists',N('Some list numbering could not be read'),()=>{
      const pair=book.pair(73);if(!pair.length)return;
      const r=book.tableReader,count=r.u16(pair.offset),start=pair.offset+2;let p=start+28*count;
      if(count>32767)B.fail('DOC_LISTS',N('The document contains too many lists'));r.check(start,count*28);
      for(let i=0;i<count;i++){
        const at=start+i*28,flags=r.u8(at+26),list={id:i,lsid:r.u32(at),template:r.u32(at+4),flags,styles:Array.from({length:9},(_,j)=>r.u16(at+8+j*2)),levels:[]};
        for(let j=0;j<(flags&1?1:9);j++){const l=level(book,r,p);p=l.end;list.levels.push(l);}book.lists.push(list);
      }
      const o=D.table(book,74,'PlfLfo');if(!o.length)return;
      const n=o.u32(0);if(n>32767)B.fail('DOC_LISTS',N('The document contains too many list overrides'));o.check(4,n*16);p=4+n*16;
      const byId=new Map(book.lists.map(l=>[l.lsid,l]));
      for(let i=0;i<n;i++){
        const at=4+i*16,list=byId.get(o.u32(at)),count=o.u8(at+12),override={id:i+1,list,levels:[]};
        o.check(p,4);p+=4;
        if(count)while(p+4<=o.length&&o.u32(p)===0xffffffff)p+=4;
        if(count>9)B.fail('DOC_LISTS',N('A list override has too many levels'));
        for(let j=0;j<count;j++){
          const start=o.i32(p),flags=o.u32(p+4);p+=8;
          const value={index:flags&15,start:flags&16?start:null};
          if(flags&32){value.level=level(book,o,p);p=value.level.end;}if(value.index<9)override.levels.push(value);
        }
        book.overrides.push(override);
      }
    },null);
  };
  D.listFor=function(book,p,where){
    if(!p.list)return null;
    // Negative ilfo is Word's "preserve indents" form of the same LFO index.
    const index=p.list<0?(p.list&65535)^65535:p.list,override=book.overrides[index-1],level=p.level||0;
    if(index===2047||!override?.list||level>8){D.notice(book,'legacy-word-list',N('List numbering could not be kept'),where,true);return null;}
    return {override,level,definition:override.levels.find(l=>l.index===level)?.level||override.list.levels[level]};
  };
})(typeof window!=='undefined'?window:globalThis);
