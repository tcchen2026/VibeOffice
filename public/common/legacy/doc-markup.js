/* Binary Word bookmarks, annotations and revision identities (MS-DOC 2.8/2.9).
 * Derived from Apache POI hwpf/model/{BookmarksTables,Sttb}.java (Apache-2.0).
 * Derived from Apache OpenOffice sw/source/filter/ww8/ww8par{,4,5}.cxx
 * (Apache-2.0): comment authors/dates and revision author/date lookup.
 * FBKF (four bytes), not the six-byte BKF used for range protection, joins
 * each bookmark start to its end. Positions count UTF-16 units in all stories. */
(function(root){
  'use strict';
  const B=root.L.legacy,D=B.doc,N=s=>s,M=D.markup={},w=(...a)=>D.w(...a);
  M.strings=function(r,state={}){
    if(!r.length)return [];
    const wide=r.u16(0)===65535;let p=wide?2:0;const count=r.u16(p),extra=r.u16(p+2),out=[];p+=4;
    for(let i=0;i<count;i++){
      const n=wide?r.u16(p):r.u8(p);p+=wide?2:1;
      const missing=n===(wide?65535:255),text=missing?'':wide?D.unicode(r,p,n*2):r.text(p,n,1252);
      if(!missing)p+=n*(wide?2:1);
      out.push({text,extra:r.reader(p,missing?0:extra)});if(!missing)p+=extra;
    }
    state.end=p;return out;
  };
  M.date=function(value){
    const minute=value&63,hour=value>>>6&31,day=value>>>11&31,month=value>>>16&15,year=1900+(value>>>20&511);
    if(!day||!month||minute>59||hour>23||month>12)return null;
    const date=new Date(Date.UTC(year,month-1,day,hour,minute));
    return date.getUTCMonth()===month-1?date.toISOString().replace('.000Z','Z'):null;
  };
  const boundary=(book,cp)=>!(/[\uD800-\uDBFF]/.test(book.text[cp-1]||'')&&/[\uDC00-\uDFFF]/.test(book.text[cp]||''));
  function ranges(book,startId,endId){
    const starts=D.table(book,startId,'Bookmark starts'),ends=D.table(book,endId,'Bookmark ends');
    if(!starts.length)return [];
    const a=D.plc(starts,4),b=D.plc(ends,0),out=[];
    for(let i=0;i<a.n;i++){const at=a.at(i),j=starts.u16(at);out.push({start:a.cp[i],end:j<b.n?b.cp[j]:null,flags:starts.u16(at+2)});}
    return out;
  }
  function event(book,record,cp,tag,attrs,reference=false){
    let at=book.markupAt.get(cp);if(!at)book.markupAt.set(cp,at=[]);
    at.push({record,tag,attrs,reference,emitted:false});
  }
  function bookmarks(book){
    if(!book.pair(22).length)return;
    const names=M.strings(D.table(book,21,'Bookmark names')),positions=ranges(book,22,23),seen=new Set();
    book.markupStats.sourceBookmarks=positions.length;
    for(let id=0;id<positions.length;id++){
      const b={...positions[id],id,name:names[id]?.text,kind:'bookmark'};
      if(b.end==null||b.end<b.start||b.end>book.text.length||!boundary(book,b.start)||!boundary(book,b.end)||!b.name||b.name.length>40||seen.has(b.name.toLowerCase())){
        D.notice(book,'legacy-word-bookmarks',N('A damaged bookmark was not kept'),'Document',true);continue;
      }
      seen.add(b.name.toLowerCase());book.bookmarks.push(b);
      const first=b.flags&127,last=(b.flags>>>8&63)-1,columns=b.flags&32768&&first<=last;
      event(book,b,b.start,'bookmarkStart',{id,name:b.name,colFirst:columns?first:null,colLast:columns?last:null});
      event(book,b,b.end,'bookmarkEnd',{id});
    }
  }
  function comments(book){
    if(!book.pair(4).length){if(book.stories.comments)D.notice(book,'legacy-word-comments',N('Comments without document positions were not kept'),'Comments',true);return;}
    const r=D.table(book,4,'Comment references'),refs=D.plc(r,30),texts=D.plc(D.table(book,5,'Comment text'),0),story=book.storyRanges.comments;
    book.markupStats.sourceComments=refs.n;
    const owners=D.optional(book,'Comment authors',N('Some comment author names could not be read'),()=>{
      const r=D.table(book,36,'Comment authors'),out=[];let p=0;
      while(p<r.length){const n=r.u16(p);p+=2;out.push(D.unicode(r,p,n*2));p+=n*2;if(out.length>B.limits.records)B.fail('DOC_COMMENTS',N('There are too many comment authors'));}return out;
    },[],false);
    const anchors=D.optional(book,'Comment ranges',N('Some comment ranges were replaced with point references'),()=>{
      const tags=M.strings(D.table(book,37,'Comment bookmark tags')),list=ranges(book,42,43),out=new Map();
      for(let i=0;i<list.length;i++)if(tags[i]?.extra.length>=6)out.set(tags[i].extra.u32(2),list[i]);return out;
    },new Map());
    const extra=D.optional(book,'Comment dates',N('Some comment dates could not be read'),()=>D.table(book,112,'Comment details'),null,false);
    for(let id=0;id<refs.n;id++){
      let cp=refs.cp[id];const where='Comment '+(id+1);
      if(cp>=book.stories.main||id>=texts.n||texts.cp[id+1]>book.stories.comments){D.notice(book,'legacy-word-comments',N('A comment with damaged text positions was not kept'),where,true);continue;}
      if(!boundary(book,cp)){cp++;D.notice(book,'legacy-word-comment-range',N('A damaged comment range was replaced with a point reference'),where,true);}
      const at=refs.at(id),initials=D.unicode(r,at+2,Math.min(9,r.u16(at))*2),tag=r.u32(at+26),anchor=anchors.get(tag);
      const b={id,kind:'comment',start:cp,end:cp,cp,body:{start:story.start+texts.cp[id],end:story.start+texts.cp[id+1]},author:owners[r.u16(at+20)]||initials||'',initials,where};
      if(anchor&&anchor.end!=null&&anchor.start<=anchor.end&&anchor.end<=book.stories.main&&boundary(book,anchor.start)&&boundary(book,anchor.end)){b.start=anchor.start;b.end=anchor.end;}
      else if(tag!==0xffffffff)D.notice(book,'legacy-word-comment-range',N('A damaged comment range was replaced with a point reference'),where,true);
      if(extra&&id*18+18<=extra.length){b.date=M.date(extra.u32(id*18));b.parent=extra.i32(id*18+10);if(extra.u32(id*18+14)&2)D.notice(book,'legacy-word-ink-comment',N('Handwriting in a comment was not kept'),where,true);}
      if(b.parent)D.notice(book,'legacy-word-comment-reply',N('Comment replies were kept as separate comments'),where,true);
      book.comments.push(b);book.commentAt.add(cp);
      if(b.start!==b.end){event(book,b,b.start,'commentRangeStart',{id});event(book,b,b.end,'commentRangeEnd',{id});}
      event(book,b,cp,'commentReference',{id},true);
    }
  }
  M.prepare=function(book){
    book.markupAt=new Map();book.bookmarks=[];book.comments=[];book.commentAt=new Set();book.revisionId=0;
    book.markupStats={sourceBookmarks:0,bookmarks:0,sourceComments:0,comments:0,insertions:0,deletions:0,formatRevisions:0};
    D.optional(book,'Bookmarks',N('Some bookmarks could not be read'),()=>bookmarks(book),null);
    D.optional(book,'Comments',N('Some comments could not be read'),()=>comments(book),null);
    book.revisionAuthors=D.optional(book,'Revision authors',N('Some revision author names could not be read'),()=>M.strings(D.table(book,51,'Revision authors')).map(x=>x.text),[],false);
  };
  M.at=function(book,range,cp){
    let xml='';
    for(const e of book.markupAt.get(cp)||[]){
      const b=e.record;if(e.emitted||b.start<range.start||b.end>range.end||b.cp!=null&&(b.cp<range.start||b.cp>=range.end))continue;
      e.emitted=true;b.emitted=(b.emitted||0)+1;
      const tag=w(e.tag,e.attrs);xml+=e.reference?w('r',{},tag):tag;
    }
    return xml;
  };
  M.writeComments=function(book,pkg){
    if(!book.comments.length)return;
    let body='';for(const c of book.comments){
      body+=w('comment',{id:c.id,author:c.author.slice(0,255),initials:c.initials,date:c.date},D.storyXML(book,c.body,c.where,{comment:true}));book.markupStats.comments++;
    }
    pkg.add('word/comments.xml',D.part('comments',body),'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml');pkg.rel('word/document.xml','comments','comments.xml');
  };
  M.finish=function(book){
    for(const b of book.bookmarks){if(b.emitted===2)book.markupStats.bookmarks++;else D.notice(book,'legacy-word-bookmarks',N('A bookmark outside the converted text was not kept'),'Bookmark '+b.name,true);}
  };
  M.revision=function(book,p,type){
    const insert=type==='ins',meta=type==='format'?p.formatRevision:{author:p[insert?'insertAuthor':'deleteAuthor'],date:p[insert?'insertDate':'deleteDate']};
    return {id:book.revisionId++,author:(book.revisionAuthors[meta?.author??0]||'').slice(0,255),date:M.date(meta?.date||0)};
  };
  M.marks=function(book,p){
    let out='';if(p.inserted){out+=w('ins',M.revision(book,p,'ins'));book.markupStats.insertions++;}if(p.deleted){out+=w('del',M.revision(book,p,'del'));book.markupStats.deletions++;}return out;
  };
  M.wrap=function(book,p,body){
    if(p.deleted){body=w('del',M.revision(book,p,'del'),body);book.markupStats.deletions++;}
    if(p.inserted){body=w('ins',M.revision(book,p,'ins'),body);book.markupStats.insertions++;}return body;
  };
  M.change=function(book,p,kind,where,base){
    if(!p.revisionFormatting)return '';
    if(!p.revisionBefore){D.notice(book,'legacy-word-revision-format',N('An earlier formatting change was not kept; the current formatting was retained'),where,true);return '';}
    book.markupStats.formatRevisions++;const r=kind===2?'rPr':'pPr';
    return w(r+'Change',M.revision(book,p,'format'),w(r,{},kind===2?D.runXML(book,p.revisionBefore,where,null,base):D.paragraphXML(book,p.revisionBefore,where)));
  };
})(typeof window!=='undefined'?window:globalThis);
