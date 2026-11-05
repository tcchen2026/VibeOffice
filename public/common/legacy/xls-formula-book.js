/* Derived from Apache OpenOffice main/sc/source/filter/excel/{xiname,xilink,xihelper}.cxx (Apache-2.0).
 * Derived from Apache POI hssf/record/{Name,ExternalName,ExternSheet,SupBook,SharedFormula,
 * Array,Table,CRN,CRNCount}Record.java (Apache-2.0). Formula ownership and links.
 * Table field order follows MS-XLS 2.4.319; OOXML assembly is independent. */
(function(root){
  'use strict';
  const B=root.L.legacy,X=B.xls,el=B.el,N=s=>s;
  const fail=(message,where)=>B.fail('FORMULA_LINK',message,where);
  const builtins=['Consolidate_Area','Auto_Open','Auto_Close','Extract','Database','Criteria','Print_Area','Print_Titles','Recorder','Data_Form','Auto_Activate','Auto_Deactivate','Sheet_Title','_FilterDatabase'];
  const key=(row,col)=>row*256+col;
  const continued=(records,i)=>{
    const record=records[i];let end=i+1,size=record.r.length;
    while(records[end]?.id===0x3c){size+=records[end++].r.length;if(size>B.limits.text)fail('Formula record is too large');}
    if(end===i+1)return record.r;
    const bytes=new Uint8Array(size);let at=0;for(let j=i;j<end;j++){bytes.set(records[j].r.bytes,at);at+=records[j].r.length;}return new B.Reader(bytes,'Continued formula record');
  };
  const url=encoded=>{
    if(!encoded||encoded==='\x00')return {kind:'deleted'};
    if(/^[\x02\x03]/.test(encoded))return {kind:'self'};
    if(/^\x01\x08.*(?:ATPVBAEN\.XLA|ANALYS32\.XLL)$/i.test(encoded))return {kind:'analysis'};
    let value='';if(encoded.charCodeAt(0)!==1)value=encoded;
    else for(let i=1;i<encoded.length;i++){
      const c=encoded.charCodeAt(i);
      if(c===1){const drive=encoded[++i];if(!drive)fail('Incomplete external workbook path');value+=drive==='@'?'//':drive+':/';}
      else if(c===2||c===3)value+='/';else if(c===4)value+='../';
      else if(c===5){const n=encoded.charCodeAt(++i);if(!Number.isFinite(n))fail('Incomplete raw workbook path');value+=encoded.slice(i+1,i+1+n);i+=n;}
      else if(c>=6&&c<=8)fail('External workbook uses an unavailable installation path');else value+=encoded[i];
    }
    value=value.replace(/\\/g,'/');if(/^[A-Za-z]:\//.test(value))value='file:///'+value;else if(value.startsWith('//'))value='file:'+value;
    if(!value||/[\x00-\x1f]/.test(value)||/^(?:javascript|vbscript|data):/i.test(value))fail('Invalid external workbook path');
    return {kind:'external',url:value.replace(/ /g,'%20')};
  };
  X.prepareFormulas=function(book,losses){
    const b8=book.version==='BIFF8',names=[],supbooks=[],xti=[],calc={},stats={source:0,converted:0,dropped:0,shared:0,array:0,dataTable:0};let current=null,cachedSheet=-1;
    const metadataLoss=(id,what,where,error,notify=false)=>{const loss=B.loss(id,what,where,'drop',String(where),notify);loss.detail={code:error.code,message:error.message};losses.push(loss);};
    const calculation=rec=>{
      const r=rec.r,id=rec.id;
      if(id===0xd)calc.calcMode=({0:'manual',1:'auto',2:'autoNoTable',65535:'autoNoTable'})[r.u16(0)];
      else if(id===0xc)calc.iterateCount=r.u16(0);
      else if(id===0x11)calc.iterate=r.u16(0)?1:null;
      else if(id===0x10){const n=r.f64(0);if(Number.isFinite(n)&&n>=0)calc.iterateDelta=n;}
      else if(id===0xe)calc.fullPrecision=r.u16(0)?null:0;
      else if(id===0xf)calc.refMode=r.u16(0)?null:'R1C1';
      else if(id===0x5f)calc.calcOnSave=r.u16(0)?null:0;
    };
    for(let i=0;i<book.globals.length;i++){
      const rec=book.globals[i],id=rec.id;if(![0x18,0x1ae,0x17,0x23,0x59,0x5a,0xd,0xc,0xe,0xf,0x10,0x11,0x5f].includes(id))continue;
      let slot;
      if(id===0x18){slot={};names.push(slot);}if(id===0x1ae){current={names:[],sheets:[],cache:new Map()};supbooks.push(current);cachedSheet=-1;}
      if(id===0x23&&current){slot={};current.names.push(slot);}
      try{
        const r=continued(book.globals,i);
        if(id===0x18){
          const flags=r.u16(0),count=r.u8(3),n=r.u16(4),scope=r.u16(8)-1;let p=14,value;
          if(b8){const wide=r.u8(p++)&1;value=wide?r.text(p,count*2,1200):B.latin16(r.sub(p,count));p+=count*(wide?2:1);}
          else{value=r.text(p,count,book.codepage);p+=count;}
          if(flags&32){const name=builtins[value.charCodeAt(0)];if(!name)fail('Unknown built-in name');value='_xlnm.'+name;}
          if(!value||/[\x00-\x20!"'\[\]:]/.test(value)||scope>=book.sheets.length)fail('Invalid defined name');
          Object.assign(slot,{name:value,scope,flags,tokens:r.sub(p,n),extra:r.sub(p+n,r.length-p-n)});
        }else if(id===0x1ae){
          if(!b8)fail('BIFF5 external workbook table');
          const count=r.u16(0),tag=r.u16(2);
          if(r.length===4&&(tag===0x401||tag===0x3a01)){current.kind=tag===0x401?'self':'addin';continue;}
          const path=X.string(r,2);Object.assign(current,url(path.value));let p=path.end;
          for(let j=0;j<count;j++){const name=X.string(r,p);current.sheets.push(name.value);p=name.end;}
        }else if(id===0x17){
          if(!b8)continue;const count=r.u16(0);r.check(2,count*6);for(let j=0;j<count;j++)xti.push({book:r.u16(2+j*6),first:r.i16(4+j*6),last:r.i16(6+j*6)});
        }else if(id===0x23&&current){
          const flags=r.u16(0),scope=r.u16(2)-1,item=X.string(r,6,b8,book.codepage,true);Object.assign(slot,{name:item.value,scope,flags});
          if(!(flags&0x1a)&&item.end+2<=r.length){const n=r.u16(item.end),p=item.end+2;slot.tokens=r.sub(p,n);slot.extra=r.sub(p+n,r.length-p-n);}
        }else if(id===0x59&&current)cachedSheet=r.u16(2);
        else if(id===0x5a&&current&&cachedSheet>=0){
          const last=r.u8(0),first=r.u8(1),row=r.u16(2),cursor={p:4};if(first>last||cachedSheet>=current.sheets.length)fail('Invalid external cache range');
          if(!current.cache.has(cachedSheet))current.cache.set(cachedSheet,new Map());const cells=current.cache.get(cachedSheet);
          for(let col=first;col<=last;col++)cells.set(key(row,col),{row,col,...X.formulaConstant(r,cursor,b8,book.codepage)});
        }else calculation(rec);
      }catch(error){
        if(error.name!=='LegacyError')throw error;if(slot)slot.invalid=true;if(id===0x1ae)current.invalid=true;
        metadataLoss(id===0x18?'legacy-defined-name':'legacy-formula-metadata',id===0x18?N('A defined name was not kept'):'Some formula reference settings could not be read',id===0x18?(slot.name||'Workbook'):rec.offset,error,id===0x18);
      }
    }
    // BIFF stores most workbook calculation options in the worksheet substream.
    // Use the first worksheet's copy; nested chart substreams do not count.
    let calculationDepth=0;
    for(const rec of book.sheets[0]?.records||[]){
      if(rec.id===0x809){calculationDepth++;continue;}if(rec.id===10){calculationDepth--;continue;}
      if(calculationDepth!==1||![0xd,0xc,0xe,0xf,0x10,0x11,0x5f].includes(rec.id))continue;
      try{calculation(rec);}catch(error){if(error.name!=='LegacyError')throw error;metadataLoss('legacy-formula-metadata','Some formula reference settings could not be read',rec.offset,error);}
    }
    let externalCount=0;for(const sb of supbooks)if(sb.kind==='external'&&!sb.invalid)sb.externalId=++externalCount;
    const context=(sheetIndex,row=0,col=0,extra={})=>({b8,cp:book.codepage,row,col,...extra,
      name:index=>{const name=names[index-1];if(!name||name.invalid)return null;const text=name.scope>=0&&name.scope!==sheetIndex?X.quoteSheet(book.sheets[name.scope].name)+'!'+name.name:name.name;return name.tokens.length?text:{text,name:text,functionOnly:true};},
      sheet:index=>{
        const ref=xti[index],sb=ref&&supbooks[ref.book];if(!sb||sb.invalid)return null;if(sb.kind==='deleted'||ref.first===-1||ref.last===-1)return '#REF!';if(!['self','external'].includes(sb.kind))return null;
        const list=sb.kind==='self'?book.sheets.map(s=>s.name):sb.sheets,a=list[ref.first],z=list[ref.last];if(a==null||z==null)return null;
        return X.quoteSheet((sb.externalId?'['+sb.externalId+']':'')+a+(ref.first!==ref.last?':'+z:''));
      },
      externalName:(index,nameIndex)=>{
        const ref=xti[index],sb=ref&&supbooks[ref.book];if(!sb||sb.invalid)return null;
        if(sb.kind==='self')return context(sheetIndex).name(nameIndex);
        const name=sb.names[nameIndex-1];if(!name||name.invalid)return null;
        if(sb.kind==='addin'||sb.kind==='analysis'&&X.addinFunctions.has(name.name.toUpperCase()))return {text:name.name,name:name.name};
        let text=null;
        if(sb.externalId){const prefix='['+sb.externalId+']';text=name.scope>=0?(sb.sheets[name.scope]!=null?X.quoteSheet(prefix+sb.sheets[name.scope])+'!'+name.name:null):prefix+'!'+name.name;}
        return text?{text,name:text}:null;
      }
    });
    // Retain names as names; do not recursively expand aliases or cyclic names.
    const defined=[];
    for(const name of names){if(name.invalid||!name.tokens.length&&!(name.flags&0xe))continue;try{
      const text=name.tokens.length?X.formula(name.tokens,name.extra,context(name.scope,0,0,{nameMode:true})):'#NAME?';
      defined.push(el('definedName',{name:name.name,localSheetId:name.scope>=0?name.scope:null,hidden:name.flags&1?1:null,function:name.flags&2?1:null,vbProcedure:name.flags&4?1:null,xlm:name.flags&8&&!(name.flags&4)?1:null},B.xml(text)));
    }catch(error){if(error.name!=='LegacyError')throw error;name.invalid=true;metadataLoss('legacy-defined-name',N('A defined name was not kept'),name.name,error,true);}}
    for(let si=0;si<book.sheets.length;si++){
      const sheet=book.sheets[si],groups=new Map(),recordIndex=new Map();let depth=0,last=null,dropped=0;const examples=[],reasons={};
      if(!sheet.cells.some(c=>c.formula))continue;
      for(let i=0;i<sheet.records.length;i++){
        const rec=sheet.records[i],id=rec.id;if(id===0x809){depth++;continue;}if(id===10){depth--;continue;}if(depth!==1)continue;
        if(id===6){recordIndex.set(rec.offset,i);last={row:rec.r.u16(0),col:rec.r.u16(2)};continue;}if(![0x4bc,0x221,0x236].includes(id))continue;
        try{
          const r=continued(sheet.records,i),range=[r.u16(0),r.u8(4),r.u16(2),r.u8(5)],kind=id===0x4bc?'shared':id===0x221?'array':'dataTable';
          if(range[0]>range[2]||range[1]>range[3])fail('Invalid formula group range');
          const anchor=kind==='shared'&&last?last:{row:range[0],col:range[1]},group={kind,range,anchor,r};
          if(kind!=='dataTable'){const p=kind==='shared'?8:12,n=r.u16(p);group.tokens=r.sub(p+2,n);group.extra=r.sub(p+2+n,r.length-p-2-n);}
          const at=key(anchor.row,anchor.col);if(groups.has(at))group.invalid=true;groups.set(at,group);
        }catch(error){if(error.name!=='LegacyError')throw error;/* Members report a local loss if no valid owner remains. */}
      }
      for(const cell of sheet.cells){if(!cell.formula)continue;stats.source++;
        try{
          let text,attrs={},member=false,kind='ordinary',tokens=cell.formula;
          if(tokens[0]===1||tokens[0]===2){
            const r=new B.Reader(tokens),group=groups.get(key(r.u16(1),r.u16(3)));if(!group||group.invalid)fail('Missing or conflicting formula group');
            const [r1,c1,r2,c2]=group.range;if(cell.row<r1||cell.row>r2||cell.col<c1||cell.col>c2)fail('Formula cell lies outside its group');kind=group.kind;
            if((tokens[0]===2)!==(kind==='dataTable'))fail('Formula group kind differs');
            if(kind==='shared')text=X.formula(group.tokens,group.extra,context(si,cell.row,cell.col,{shared:true}));
            else{
              member=cell.row!==r1||cell.col!==c1;attrs={t:kind,ref:X.range(...group.range)};
              if(kind==='array'){
                if(group.text==null)group.text=X.formula(group.tokens,group.extra,context(si,r1,c1));text=group.text;attrs.aca=group.r.u16(6)&1?1:null;
              }else{
                const r=group.r,flags=r.u16(6),two=!!(flags&8);attrs.dt2D=two?1:null;attrs.dtr=!two&&flags&4?1:null;attrs.del1=flags&16?1:null;attrs.del2=two&&flags&32?1:null;attrs.ca=flags&1?1:null;
                if(!attrs.del1){if(r.u16(10)>255)fail('Invalid table input cell');attrs.r1=X.address(r.u16(8),r.u16(10));}
                if(two&&!attrs.del2){if(r.u16(14)>255)fail('Invalid table second input cell');attrs.r2=X.address(r.u16(12),r.u16(14));}text='';
              }
            }
          }else{
            const i=recordIndex.get(cell.offset),r=i==null?cell.formulaRecord.r:continued(sheet.records,i),p=22+tokens.length;
            text=X.formula(tokens,r.sub(p,r.length-p),context(si,cell.row,cell.col));attrs.ca=r.u16(14)&1?1:null;
          }
          cell.formulaOut={text,attrs,member};stats.converted++;if(kind!=='ordinary')stats[kind]++;
        }catch(error){
          if(error.name!=='LegacyError')throw error;dropped++;stats.dropped++;if(examples.length<12)examples.push(X.address(cell.row,cell.col));const why=error.message+(error.where!=null?' ('+error.where+')':'');reasons[why]=(reasons[why]||0)+1;
        }
      }
      if(dropped){const loss=B.loss('legacy-formulas',N('Some formulas were replaced with their saved results'),sheet.name,'drop',sheet.name,true);loss.detail={cells:dropped,examples,reasons};losses.push(loss);}
    }
    return {stats,defined:defined.length?el('definedNames',{},defined.join('')):'',calc,
      addExternalParts(pkg){
        const refs=[];
        for(const sb of supbooks){if(!sb.externalId)continue;const name='xl/externalLinks/externalLink'+sb.externalId+'.xml';
          const sheets=el('sheetNames',{},sb.sheets.map(s=>el('sheetName',{val:s})).join('')),data=[];
          for(const [sheetId,cells]of sb.cache){const rows=[];let row=-1,body=[];for(const c of [...cells.values()].sort((a,b)=>a.row-b.row||a.col-b.col)){
            if(c.row!==row){if(row>=0)rows.push(el('row',{r:row+1},body.join('')));row=c.row;body=[];}
            body.push(el('cell',{r:X.address(c.row,c.col),t:c.type==='n'||c.type==='blank'?null:c.type},c.type==='blank'?undefined:el('v',{},B.xstring(c.type==='b'?c.value?'1':'0':c.value))));
          }if(row>=0)rows.push(el('row',{r:row+1},body.join('')));data.push(el('sheetData',{sheetId},rows.join('')));}
          const extNames=[];
          for(const n of sb.names){if(n.invalid||!n.tokens?.length)continue;try{
            const ctx=context(-1,0,0,{nameMode:true,externalDefinition:true});ctx.sheet=(first,last)=>sb.sheets[first]!=null&&sb.sheets[last]!=null?X.quoteSheet(sb.sheets[first]+(first!==last?':'+sb.sheets[last]:'')):null;
            extNames.push(el('definedName',{name:n.name,refersTo:X.formula(n.tokens,n.extra,ctx),sheetId:n.scope>=0?n.scope:null}));
          }catch(error){if(error.name!=='LegacyError')throw error;metadataLoss('legacy-external-name','An external workbook name could not be read',n.name,error);}}
          const rid=pkg.rel(name,'externalLinkPath',sb.url,{external:true});
          pkg.add(name,B.xmlDeclaration+el('externalLink',{xmlns:B.ns.s,'xmlns:r':B.ns.officeRel.slice(0,-1)},el('externalBook',{'r:id':rid},sheets+(extNames.length?el('definedNames',{},extNames.join('')):'')+(data.length?el('sheetDataSet',{},data.join('')):''))),'application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml');
          refs.push(el('externalReference',{'r:id':pkg.rel('xl/workbook.xml','externalLink','externalLinks/externalLink'+sb.externalId+'.xml')}));
        }return refs.length?el('externalReferences',{},refs.join('')):'';
      }
    };
  };
})(typeof window!=='undefined'?window:globalThis);
