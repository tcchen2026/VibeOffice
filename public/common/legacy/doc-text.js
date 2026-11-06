/* Word stories -> paragraphs and runs. Binary marker meanings: MS-DOC §2.4.
 * Complex fields keep their instructions and cached results separately.
 * Tables/revisions and drawings retain local fallbacks until their writers run. */
(function(root){
  'use strict';
  const B=root.L.legacy,D=B.doc,N=s=>s,w=(...args)=>D.w(...args);
  const fallbackChar=book=>({props:book.defaults});
  D.storyXML=function(book,range,where,options={}){
    const nodes=[],fields=[],text=book.text;let start=range.start,paragraphs=0;
    const sectionEnds=options.main?new Map(book.sections.slice(0,-1).map(s=>[s.end,s])):new Map();
    function paragraph(a,z,mark){
      const xml=new B.XML();
      const place=where+', paragraph '+(++paragraphs),markCp=Math.max(a,z-1);
      const paragraph=D.optional(book,place,N('Some paragraph formatting could not be read'),()=>D.properties(book,markCp,1,null,place),{props:{},style:book.styles[0]}),p=paragraph.props;
      D.formatWarnings(book,p,place);let ppr=D.paragraphXML(book,p,place,{style:paragraph.style,explicitList:true});
      const markValue=D.optional(book,place,N('Some text formatting could not be read'),()=>D.properties(book,markCp,2,paragraph.style,place),fallbackChar(book)),markProps=markValue.props;
      // The paragraph mark's font affects blank-line height as well as editing.
      ppr+=w('rPr',{},(D.tables.rowEnd(p,mark)?'':D.markup.marks(book,markProps))+D.runXML(book,markProps,place,null,markValue.base)+D.markup.change(book,markProps,2,place,markValue.base));
      const section=sectionEnds.get(z);if(section)ppr+=D.sectionXML(book,section);
      ppr+=D.markup.change(book,p,1,place);
      const prefix=w('pPr',{},ppr);xml.start('w:p').append(prefix);
      let cp=a,end=mark?z-1:z;
      while(cp<end){
        const value=D.optional(book,place,N('Some text formatting could not be read'),()=>D.properties(book,cp,2,paragraph.style,place),{...fallbackChar(book),end});
        const c=value.props;let stop=Math.max(cp+1,Math.min(end,value.end||end));
        // Formatting can split a surrogate pair; keep the glyph with its first
        // code unit rather than replacing both halves in separate XML nodes.
        if(stop<end&&/[\uD800-\uDBFF]/.test(text[stop-1])&&/[\uDC00-\uDFFF]/.test(text[stop]))stop++;
        const chunks=[];let ordinary='';
        const flush=()=>{if(ordinary){const instruction=fields.length&&!fields.at(-1).result;chunks.push(B.el(instruction?(c.deleted?'w:delInstrText':'w:instrText'):(c.deleted?'w:delText':'w:t'),{'xml:space':'preserve'},B.xml(ordinary)));book.textStats[instruction?'instructionCharacters':c.deleted?'deletedCharacters':'characters']+=ordinary.length;ordinary='';}};
        const emit=()=>{flush();if(chunks.length){D.formatWarnings(book,c,place);const props=D.runXML(book,c,place,null,value.base)+D.markup.change(book,c,2,place,value.base);xml.append(D.markup.wrap(book,c,w('r',{},(props?w('rPr',{},props):'')+chunks.join(''))));book.textStats.runs++;chunks.length=0;}};
        for(let i=cp;i<stop;i++){
          const markup=D.markup.at(book,range,i);if(markup){emit();xml.append(markup);}
          const ch=text.charCodeAt(i);
          if(ch===19){flush();if(fields.length>=B.limits.depth)B.fail('DOC_FIELD_DEPTH',N('Fields in this document are nested too deeply'));const f=D.fields.begin(book,range,i,where,fields);fields.push(f);if(f.keep)chunks.push(w('fldChar',{fldCharType:'begin',fldLock:f.node.flags&16?1:null,dirty:f.node.flags&4?1:null},f.node.formXML||undefined));continue;}
          if(ch===20){flush();const f=fields.at(-1);if(f){if(f.keep&&!f.result)chunks.push(w('fldChar',{fldCharType:'separate'}));f.result=true;}continue;}
          if(ch===21){flush();const f=fields.pop();if(f?.keep)chunks.push(w('fldChar',{fldCharType:'end'}));continue;}
          if(fields.some(f=>!f.keep&&!f.result))continue;
          const note=options.main?book.noteAt.get(i):null;
          if(note){flush();chunks.push(w(note.kind+'Reference',{id:note.id,customMarkFollows:note.automatic?null:1}));if(note.automatic&&ch===2)continue;}
          if(ch===2){flush();if(options.note)chunks.push(w(options.note+'Ref'));else if(!note)D.notice(book,'legacy-word-note-reference',N('A note reference without a matching note was not kept'),where,true);continue;}
          if(ch===1||ch===8){flush();if(ch===1&&fields.at(-1)?.node?.formMark===i)continue;D.notice(book,'legacy-word-picture',N('Pictures and drawing objects were not kept'),where,true);book.textStats.objects++;continue;}
          if(ch===5){flush();if(options.comment)chunks.push(w('annotationRef'));else if(!book.commentAt.has(i))D.notice(book,'legacy-word-comments',N('A comment reference without a matching comment was not kept'),where,true);continue;}
          if(ch===9){flush();chunks.push(w('tab'));continue;}
          if(ch===11||ch===12||ch===14){flush();chunks.push(w('br',{type:ch===12?'page':ch===14?'column':null,clear:ch===11?['none','left','right','all'][c.breakClear]:null}));continue;}
          if(ch===3||ch===4){flush();if(options.separator)chunks.push(w(ch===3?'separator':'continuationSeparator'));continue;}
          if(ch===30||ch===31){flush();chunks.push(w(ch===30?'noBreakHyphen':'softHyphen'));continue;}
          if(ch<32&&ch!==10){D.notice(book,'legacy-word-control',N('An unsupported text control was not kept'),where,true,'0x'+ch.toString(16));continue;}
          if(ch===10){flush();chunks.push(w('br'));continue;}
          if(c.symbol){flush();const font=book.fonts[c.symbol.font]?.name;if(font)chunks.push(w('sym',{font,char:(c.symbol.character|0xf000).toString(16).padStart(4,'0').toUpperCase()}));else ordinary+=String.fromCharCode(c.symbol.character);continue;}
          ordinary+=text[i];
        }
        emit();
        cp=stop;
      }
      xml.append(D.markup.at(book,range,end));if(z===range.end&&end!==z)xml.append(D.markup.at(book,range,z));
      const note=options.main&&mark?book.noteAt.get(end):null;
      if(note)xml.append(w('r',{},w(note.kind+'Reference',{id:note.id,customMarkFollows:note.automatic?null:1})));
      xml.end();book.textStats.paragraphs++;
      const result=xml.finish().slice(B.xmlDeclaration.length);nodes.push({a,z,mark,p,style:paragraph.style,markProps,xml:result,inner:result.slice(5+prefix.length,-6)});
    }
    for(let cp=range.start;cp<range.end;cp++){
      const ch=text[cp];if(ch==='\r'||ch==='\x07'||ch==='\x0c'&&sectionEnds.has(cp+1)){paragraph(start,cp+1,ch);start=cp+1;}
    }
    if(start<range.end)paragraph(start,range.end,null);else if(!paragraphs)paragraph(start,start,null);
    // XML.finish also bounds total output; the story is an XML fragment here.
    return new B.XML().append(D.tables.story(book,nodes,where)).finish().slice(B.xmlDeclaration.length);
  };
  D.writeStories=function(book,pkg){
    let header=0,footer=0;const main='word/document.xml',ct='application/vnd.openxmlformats-officedocument.wordprocessingml.',headers=new Map(book.headers.map(h=>[h.index,h]));
    for(let i=0;i<book.sections.length;i++){
      const section=book.sections[i];section.references='';
      for(const slot of [0,1,4,2,3,5]){
        const range=headers.get(6+i*6+slot);if(!range?.present)continue;
        const kind=[0,1,4].includes(slot)?'header':'footer',type=slot===0||slot===2?'even':slot>=4?'first':'default';
        const index=kind==='header'?++header:++footer,name=kind+index+'.xml',where=(kind==='header'?'Header':'Footer')+' '+index;
        const content=D.storyXML(book,range,where);pkg.add('word/'+name,D.part(kind==='header'?'hdr':'ftr',content),ct+kind+'+xml');
        section.references+=w(kind+'Reference',{type,'r:id':pkg.rel(main,kind,name)});
      }
    }
    for(const [kind,list,base] of [['footnote',book.footnotes,0],['endnote',book.endnotes,3]]){
      if(!list.length)continue;let body='';
      for(let i=0;i<2;i++){
        const range=headers.get(base+i),type=i?'continuationSeparator':'separator';
        const content=range?.present?D.storyXML(book,range,kind+' '+type,{separator:true}):w('p',{},w('r',{},w(type)));
        body+=w(kind,{type,id:i-1},content);
      }
      const notice=headers.get(base+2);if(notice?.present)body+=w(kind,{type:'continuationNotice',id:list.length+1},D.storyXML(book,notice,kind+' continuation notice',{separator:true}));
      for(const note of list)body+=w(kind,{id:note.id},D.storyXML(book,note,note.where,{note:kind}));
      pkg.add('word/'+kind+'s.xml',D.part(kind+'s',body),ct+kind+'s+xml');pkg.rel(main,kind+'s',kind+'s.xml');
    }
    book.textStats.headers=header;book.textStats.footers=footer;
    return D.storyXML(book,book.storyRanges.main,'Document',{main:true})+D.sectionXML(book,book.sections.at(-1));
  };
})(typeof window!=='undefined'?window:globalThis);
