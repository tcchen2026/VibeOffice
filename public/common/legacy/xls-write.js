/* Binary Excel workbook -> SpreadsheetML: cells, formatting, views and formulas.
 * Derived from Apache OpenOffice main/sc/source/filter/excel/ (Apache-2.0).
 * Derived from Apache POI hssf/record/common/{ExtRst,PhRun}.java (Apache-2.0).
 * Drawings/charts and later features are explicit losses until
 * their converter phases are implemented. Documents stay entirely in memory. */
(function(root){
  'use strict';
  const B=root.L.legacy,X=B.xls,el=B.el,N=s=>s;
  const types={workbook:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',template:'application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml',sheet:'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',styles:'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',strings:'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml'};
  const errors={0:'#NULL!',7:'#DIV/0!',15:'#VALUE!',23:'#REF!',29:'#NAME?',36:'#NUM!',42:'#N/A',43:'#GETTING_DATA'};
  const text=value=>el('t',/^\s|\s$/.test(value)?{'xml:space':'preserve'}:{},B.xstring(value));
  X.cellText=function(value,losses,where){
    if(value.length>32767){value=value.slice(0,32767).replace(/[\ud800-\udbff]$/,'');losses.push(B.loss('legacy-cell-text-limit',N('Text beyond the 32,767-character cell limit was not kept'),where,'drop',where,true));}
    return value;
  };
  X.richText=function(item,styles,losses,where){
    const value=X.cellText(item.value||'',losses,where);
    const runs=(item.runs||[]).filter(r=>r.at>=0&&r.at<value.length).sort((a,b)=>a.at-b.at);let out='';
    if(!runs.length)out=text(value);
    else{
      if(runs[0].at>0)out+=el('r',{},text(value.slice(0,runs[0].at)));
      for(let i=0;i<runs.length;i++){const end=runs[i+1]?.at??value.length;if(end>runs[i].at)out+=el('r',{},el('rPr',{},styles.fontBody(runs[i].font))+text(value.slice(runs[i].at,end)));}
    }
    if(item.extension?.length){
      try{
        const r=new B.Reader(item.extension,'SST phonetic text');if(r.u16(0)===0xffff)return out;
        if(r.u16(0)!==1)B.fail('PHONETIC_HEADER',N('The cell pronunciation text could not be read'));
        const end=r.u16(2)+4,len=r.u16(10),other=r.u16(12);r.check(0,end);if(len&&other!==len)B.fail('PHONETIC_LENGTH',N('The cell pronunciation text has an invalid length'));
        const phonetic=r.text(14,len*2,1200),p=14+len*2,count=Math.floor((end-p)/6),pr=[];r.check(p,count*6);
        for(let i=0;i<count;i++){const at=p+i*6,start=r.u16(at),next=i+1<count?r.u16(at+6):len,sb=r.u16(at+2),eb=sb+r.u16(at+4);if(start>next||next>len||eb>value.length||sb>=eb)B.fail('PHONETIC_RANGE',N('The cell pronunciation text has an invalid range'));pr.push(el('rPh',{sb,eb},text(phonetic.slice(start,next))));}
        if(pr.length){const options=r.u16(6);out+=pr.join('')+el('phoneticPr',{fontId:styles.fontIndex(r.u16(4)),type:['halfwidthKatakana','fullwidthKatakana','Hiragana','noConversion'][options&3],alignment:['noControl','left','center','distributed'][options>>>2&3]});}
      }catch(e){if(e.name!=='LegacyError')throw e;losses.push(B.loss('legacy-phonetic-text',N('Cell pronunciation text was not kept'),where,'drop',where,true));}
    }
    return out;
  };
  X.convert=async function(bytes,options={}){
    const header=new B.Reader(bytes,'Excel file');if(header.length>=2&&[9,0x209,0x409,0x809].includes(header.u16(0)))B.fail('XLS_VERSION',N('This Excel file format is not supported yet'));
    const cfb=B.cfb(bytes),stream=cfb.stream('Workbook')||cfb.stream('Book');if(!stream)B.fail('XLS_STREAM',N('This file has no Excel workbook stream'));
    const book=X.read(stream,{rich:true}),losses=[];
    if(book.encrypted)B.fail(options.password?'ENCRYPTION_UNSUPPORTED':'PASSWORD_REQUIRED',options.password?N('Encrypted binary workbooks cannot be opened yet'):N('This workbook needs a password'));
    if(!['BIFF8','BIFF5/7'].includes(book.version))B.fail('XLS_VERSION',N('This older Excel file format is not supported yet'),book.version);
    if(!book.sheets.length)B.fail('XLS_SHEETS',N('The workbook has no worksheets'));
    const names=new Set();for(const sheet of book.sheets){if(!sheet.name||sheet.name.length>31||/[\[\]:*?\\/\x00-\x1f]/.test(sheet.name)||names.has(sheet.name.toLowerCase()))B.fail('SHEET_NAME',N('The workbook contains an invalid or duplicate sheet name'),sheet.name);names.add(sheet.name.toLowerCase());if(sheet.kind!==0)B.fail('SHEET_KIND',N('This workbook contains a sheet type that cannot be converted yet'),sheet.name);}
    const formulas=X.prepareFormulas(book,losses);
    book.theme=await X.theme(book,losses);
    const styles=X.styles(book,losses),pkg=new B.Package(),strings=[],stringMap=new Map();let stringUses=0,currentSheet='';
    const sstIndex=(item,where)=>{const xml=el('si',{},X.richText(item,styles,losses,where));if(!stringMap.has(xml)){stringMap.set(xml,strings.length);strings.push(xml);}stringUses++;return stringMap.get(xml);};
    const sourceSST=new Map();
    const cellXml=c=>{
      const attrs={r:X.address(c.row,c.col),s:styles.xfMap.get(c.xf)||null},f=c.formulaOut;let value='';
      if(f&&(c.type==='s'||c.type==='empty')){attrs.t='str';value=B.xstring(X.cellText(String(c.value??''),losses,currentSheet+'!'+attrs.r));}
      else if(c.type==='s'||c.type==='empty'){
        const where=currentSheet+'!'+attrs.r;let index;if(c.sst!=null){if(!sourceSST.has(c.sst))sourceSST.set(c.sst,sstIndex(book.strings[c.sst],where));else stringUses++;index=sourceSST.get(c.sst);}
        else index=sstIndex(c.rich||{value:String(c.value??'')},where);attrs.t='s';value=String(index);
      }else if(c.type==='n'){if(!Number.isFinite(c.value))B.fail('CELL_NUMBER',N('A cell contains an invalid numeric value'),attrs.r);value=String(c.value);}
      else if(c.type==='b'){attrs.t='b';value=c.value?'1':'0';}
      else if(c.type==='e'){attrs.t='e';value=errors[c.value];if(!value)B.fail('CELL_ERROR',N('A cell contains an invalid error value'),attrs.r);}
      return el('c',attrs,c.type==='blank'?undefined:(f&&!f.member?el('f',f.attrs,B.xml(f.text)):'')+el('v',{},attrs.t==='str'?value:B.xml(value)));
    };
    const notice=(sheet,id,message,where=sheet.name)=>losses.push(B.loss(id,message,where,'drop',sheet.name,true));
    let active=0,first=0,windowAttrs={},date1904=false,template=false,totalCells=0;
    for(const rec of book.globals){const r=rec.r;
      if(rec.id===0x22)date1904=!!r.u16(0);else if(rec.id===0x60)template=true;
      else if(rec.id===0x3d&&r.length>=18){active=r.u16(10);first=r.u16(12);const flags=r.u16(8);windowAttrs={xWindow:r.u16(0),yWindow:r.u16(2),windowWidth:r.u16(4),windowHeight:r.u16(6),showHorizontalScroll:flags&8?null:0,showVerticalScroll:flags&16?null:0,showSheetTabs:flags&32?null:0,tabRatio:r.u16(16)};}
    }
    if(!book.sheets.some(s=>!(s.state&3))){book.sheets[0].state=0;losses.push(B.loss('legacy-visible-sheet',N('The first worksheet was made visible'),book.sheets[0].name,'convert'));}
    active=Math.min(active,book.sheets.length-1);if(book.sheets[active].state&3)active=book.sheets.findIndex(s=>!(s.state&3));first=Math.min(first,book.sheets.length-1);
    const sheetList=[];
    for(let i=0;i<book.sheets.length;i++){
      const sheet=book.sheets[i],meta=X.sheetMetadata(sheet,losses);currentSheet=sheet.name;if(meta.dialog)B.fail('SHEET_KIND',N('Excel dialog sheets cannot be converted yet'),sheet.name);
      const seenCells=new Map();for(const c of sheet.cells)seenCells.set(c.row*256+c.col,c);const cells=[...seenCells.values()].sort((a,b)=>a.row-b.row||a.col-b.col);totalCells+=cells.length;
      const rowNumbers=new Set(meta.rows.keys());for(const c of cells)rowNumbers.add(c.row);const rowList=[...rowNumbers].sort((a,b)=>a-b),data=[];let at=0;
      for(const row of rowList){const stored=meta.rows.get(row),attrs=stored?{...stored}:{r:row+1};if(attrs.sourceXF!=null){attrs.s=styles.xfMap.get(attrs.sourceXF)||0;attrs.customFormat=1;}if(attrs.defaultHeight||attrs.ht<=0)delete attrs.ht;delete attrs.sourceXF;delete attrs.defaultHeight;const body=[];while(at<cells.length&&cells[at].row===row)body.push(cellXml(cells[at++]));data.push(el('row',attrs,body.length?body.join(''):undefined));}
      const view={workbookViewId:0,...meta.view};if(i===active)view.tabSelected=1;
      let selections=meta.selections;if(!selections.length)selections=[{activeCell:'A1',sqref:'A1'}];
      if(!meta.pane)selections=selections.filter(s=>!s.pane||s.pane==='topLeft').map(({pane,...s})=>s);
      const viewXml=el('sheetViews',{},el('sheetView',view,(meta.pane?el('pane',meta.pane):'')+selections.map(s=>el('selection',s)).join('')));
      const cols=meta.cols.map(({sourceXF,...c})=>el('col',{...c,style:styles.xfMap.get(sourceXF)||null}));
      const minRow=cells.length?cells[0].row:0,maxRow=cells.length?cells.at(-1).row:0,minCol=cells.reduce((n,c)=>Math.min(n,c.col),255),maxCol=cells.reduce((n,c)=>Math.max(n,c.col),0);
      const dimension=cells.length?X.range(minRow,minCol,maxRow,maxCol):'A1';
      const xml=B.xmlDeclaration+el('worksheet',{xmlns:B.ns.s,'xmlns:r':B.ns.officeRel.slice(0,-1)},el('sheetPr',{},el('outlinePr',meta.outline))+el('dimension',{ref:dimension})+viewXml+el('sheetFormatPr',meta.format)+(cols.length?el('cols',{},cols.join('')):'')+el('sheetData',{},data.join(''))+(meta.merges.length?el('mergeCells',{count:meta.merges.length},meta.merges.map(m=>el('mergeCell',{ref:X.range(...m)})).join('')):''));
      const name='xl/worksheets/sheet'+(i+1)+'.xml';pkg.add(name,xml,types.sheet);const rid=pkg.rel('xl/workbook.xml','worksheet','worksheets/sheet'+(i+1)+'.xml');sheetList.push(el('sheet',{name:sheet.name,sheetId:i+1,'r:id':rid,state:sheet.state&2?'veryHidden':sheet.state&1?'hidden':null}));
      const ids=new Set(meta.records.map(r=>r.id));
      if(ids.has(0x1c))notice(sheet,'legacy-comments',N('Cell comments were not kept'));
      if(meta.records.some(r=>r.id===0x5d&&r.r.length>=6&&![0,19].includes(r.r.u16(4))))notice(sheet,'legacy-drawings',N('Drawing objects were not kept'));
      if(ids.has(0x1b8))notice(sheet,'legacy-hyperlinks',N('Cell hyperlinks were not kept'));
      if(ids.has(0x1be))notice(sheet,'legacy-validation',N('Data validation rules were not kept'));
      if(ids.has(0x1b1)||ids.has(0x87a))notice(sheet,'legacy-conditional-formatting',N('Conditional formatting rules were not kept'));
      if(ids.has(0x9d))notice(sheet,'legacy-filters',N('AutoFilter settings were not kept'));
      if(ids.has(0xb0))notice(sheet,'legacy-pivot-tables',N('PivotTables were replaced with their saved cells'));
      if(meta.records.some(r=>[0x12,0xdd].includes(r.id)&&r.r.length>=2&&r.r.u16(0)))notice(sheet,'legacy-protection',N('Worksheet protection was not kept'));
      if(ids.has(0xa1))losses.push(B.loss('legacy-print-settings',N('Some print settings were not kept'),sheet.name,'drop',sheet.name));
    }
    if(cfb.entries.some(e=>e.type===2&&/(?:^|\/)VBA\/(?:dir|_VBA_PROJECT)$/i.test(e.path)))losses.push(B.loss('legacy-macros',N('Macros were not kept'),'VBA','drop','Workbook',true));
    if(book.globals.some(r=>[0x12,0x19].includes(r.id)&&r.r.length>=2&&r.r.u16(0)))losses.push(B.loss('legacy-workbook-protection',N('Workbook protection was not kept'),'Workbook','drop','Workbook',true));
    pkg.add('xl/styles.xml',styles.xml,types.styles);pkg.rel('xl/workbook.xml','styles','styles.xml');
    if(book.theme){pkg.add('xl/theme/theme1.xml',book.theme,'application/vnd.openxmlformats-officedocument.theme+xml');pkg.rel('xl/workbook.xml','theme','theme/theme1.xml');}
    if(strings.length){pkg.add('xl/sharedStrings.xml',B.xmlDeclaration+el('sst',{xmlns:B.ns.s,count:stringUses,uniqueCount:strings.length},strings.join('')),types.strings);pkg.rel('xl/workbook.xml','sharedStrings','sharedStrings.xml');}
    const type=template?'xltx':'xlsx',external=formulas.addExternalParts(pkg);
    pkg.add('xl/workbook.xml',B.xmlDeclaration+el('workbook',{xmlns:B.ns.s,'xmlns:r':B.ns.officeRel.slice(0,-1)},el('workbookPr',{date1904:date1904?1:null})+el('bookViews',{},el('workbookView',{...windowAttrs,activeTab:active,firstSheet:first}))+el('sheets',{},sheetList.join(''))+external+formulas.defined+el('calcPr',formulas.calc)),template?types.template:types.workbook);
    pkg.rel('','officeDocument','xl/workbook.xml');losses.push(...B.addProperties(pkg,cfb));
    return {pkg:await pkg.finish(),type,losses,stats:{sheets:book.sheets.length,cells:totalCells,strings:strings.length,formulas:formulas.stats}};
  };
})(typeof window!=='undefined'?window:globalThis);
