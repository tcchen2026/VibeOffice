/* Native WordprocessingML tables from row-ending TAPs and paragraph depth.
 * Derived from Apache OpenOffice sw/source/filter/ww8/{ww8par2,ww8par6}.cxx
 * (Apache-2.0): table bands, floating anchors, cell/row boundaries and borders.
 * MS-DOC 2.4.3/2.4.4; ECMA-376 17.4. Each paragraph is consumed once.
 * Derived from Open XML SDK data/schemas/schemas_openxmlformats_org_wordprocessingml_2006_main.json (MIT): property order. */
(function(root){
  'use strict';
  const B=root.L.legacy,D=B.doc,T=D.tables,N=s=>s,w=(...a)=>D.w(...a),sides=['top','left','bottom','right','insideH','insideV'];
  const val=(name,value)=>value==null?'':w(name,{val:value});
  T.depth=p=>p.inTable?Math.max(1,p.tableDepth||1):0;
  T.rowEnd=(p,mark)=>T.depth(p)===1?mark==='\x07'&&p.tableEnd:T.depth(p)>1&&p.innerRow;
  const cellEnd=n=>T.depth(n.p)===1?n.mark==='\x07':n.p.innerCell;
  function width(book,name,p,where,fallback){
    if(!p)return fallback==null?'':w(name,{w:Math.round(fallback),type:'dxa'});
    const type=['nil','auto','pct','dxa'][p.type];if(!type)return width(book,name,null,where,fallback);
    return w(name,{w:p.type<2?0:D.bound(book,p.value,name==='tblInd'?-31680:0,31680,where),type});
  }
  function margins(book,name,p,where){return w(name,{},sides.slice(0,4).filter(s=>p[s]!=null).map(s=>width(book,s,{type:3,value:p[s]},where)).join(''));}
  function borders(book,name,p,where){const body=sides.filter(s=>p[s]).map(s=>D.borderXML(book,s,p[s],where)).join('');return body?w(name,{},body):'';}
  function position(p){
    if(!p)return null;const vertical=p.position==null?0:p.position>>>4&3,horizontal=p.position==null?0:p.position>>>6&3;
    if(!p.y||vertical===3||horizontal===3)return null;
    const xa=({0:'left',[-4]:'center',[-8]:'right',[-12]:'inside',[-16]:'outside'})[p.x||0],ya=({[-4]:'top',[-8]:'center',[-12]:'bottom',[-16]:'inside',[-20]:'outside'})[p.y];
    return {leftFromText:p.textLeft||0,rightFromText:p.textRight||0,topFromText:p.textTop||0,bottomFromText:p.textBottom||0,vertAnchor:['margin','page','text'][vertical],horzAnchor:['text','margin','page'][horizontal],tblpXSpec:xa,tblpX:xa?null:(p.x||1)-1,tblpYSpec:ya,tblpY:ya?null:p.y-1};
  }
  function properties(book,p,where,indent,exception=false){
    let out='';
    if(!exception){
      const anchor=position(p);if(anchor)out+=w('tblpPr',anchor);
      if(p.noOverlap)out+=val('tblOverlap','never');if(p.rtl)out+=w('bidiVisual');
    }
    out+=width(book,'tblW',p.width||{type:1,value:0},where);
    let jc=p.jc??0;if(p.rtl&&p.logical&&jc<3)jc=2-jc;out+=val('jc',['left','center','right'][jc]);
    if(p.spacing!=null)out+=width(book,'tblCellSpacing',{type:3,value:p.spacing},where);
    out+=width(book,'tblInd',p.indent,where,indent)+borders(book,'tblBorders',p.borders,where);
    if(p.shading)out+=D.shadingXML(book,p.shading,where);
    out+=w('tblLayout',{type:p.autoFit?'autofit':'fixed'})+margins(book,'tblCellMar',p.padding,where);
    // No style is invented from TLP.itl: its old auto-format has already been
    // materialized as row and cell properties in the binary file.
    return out;
  }
  function rowProperties(book,p,where,before,after,grid){
    let out=before?val('gridBefore',before):'';if(after)out+=val('gridAfter',after);
    if(before)out+=width(book,'wBefore',p.before,where,grid[before]-grid[0]);
    if(after)out+=width(book,'wAfter',p.after,where,grid.at(-1)-grid[grid.length-1-after]);
    if(p.height)out+=w('trHeight',{val:Math.min(31680,Math.abs(p.height)),hRule:p.height<0?'exact':'atLeast'});
    if(p.cantSplit??p.cantSplitOld)out+=w('cantSplit');if(p.header)out+=w('tblHeader');return out;
  }
  function cellProperties(book,c,where,span,size,vertical){
    let out=width(book,'tcW',c.width?.type>=2?c.width:null,where,size);
    if(span>1)out+=val('gridSpan',span);
    // Office 97's first/continuation bit flags are 1/2 in the source files,
    // as read by AOO and POI. Keep the original cells (including hidden text)
    // with native hMerge, so unmerging does not lose their contents.
    if(c.hmerge)out+=val('hMerge',c.hmerge&1?'restart':'continue');
    if(vertical)out+=val('vMerge',vertical);
    out+=borders(book,'tcBorders',c.borders,where);if(c.shading)out+=D.shadingXML(book,c.shading,where);
    if(c.noWrap)out+=w('noWrap');if(Object.keys(c.padding).length)out+=margins(book,'tcMar',c.padding,where);
    if(c.flow)out+=val('textDirection',({1:'tbRl',3:'btLr',4:'lrTbV',5:'tbRlV'})[c.flow]);
    if(c.fit)out+=w('tcFitText');if(c.align)out+=val('vAlign',['top','center','bottom'][c.align]);if(c.hideMark)out+=w('hideMark');return out;
  }
  function tail(body,extra){
    if(!body.endsWith('</w:p>'))body+=w('p',{},'');
    return extra?body.slice(0,-6)+extra+'</w:p>':body;
  }
  function write(book,table){
    const where=table.where,rows=table.rows,props=[];
    for(const row of rows){
      const p=row.mark?D.optional(book,where,N('Some table formatting could not be read'),()=>T.properties(book,row.mark,where),null):null;props.push(p);
    }
    // Adjacent rows at the same nesting depth can belong to different floating
    // tables. As AOO's TestSameApo does, end the table when its anchor changes.
    // Nested tables cannot float independently of their containing cell.
    let start=0,anchor=JSON.stringify(position(props[0])),out='';
    for(let i=1;i<=rows.length;i++){
      const next=i<rows.length?JSON.stringify(position(props[i])):null;
      if(i===rows.length||table.depth===1&&next!==anchor){
        out+=writeRows(book,{...table,rows:rows.slice(start,i)},props.slice(start,i));start=i;anchor=next;
      }
    }
    return out;
  }
  function writeRows(book,table,props){
    const where=table.where,rows=table.rows;
    const blocks=list=>list.map(n=>n.kind==='table'?write(book,n):n.xml).join('');
    const fallback=()=>{
      D.notice(book,'legacy-word-tables',N('A damaged table was replaced with its cell text'),where,true);book.tableStats.tableFallbacks++;
      return rows.map(row=>row.cells.map(blocks).join('')+(row.mark?.inner?w('p',{},row.mark.inner):'')).join('');
    };
    if(!rows.length||props.some((p,i)=>rows[i].damaged||!p||!p.cells.length||p.cells.length!==rows[i].cells.length))return fallback();
    const grid=[...new Set(props.flatMap(p=>p.edges))].sort((a,b)=>a-b);if(grid.length<2||grid.length>4096)return fallback();
    const indices=new Map(grid.map((x,i)=>[x,i])),indent=grid[0]+props[0].padding.left;
    const previous=p=>p.revisionFormatting&&p.revisionBefore;
    const change=(p,name,inner)=>{book.markupStats.formatRevisions++;return w(name+'Change',D.markup.revision(book,p,'format'),w(name,{},inner));};
    const oldFirst=previous(props[0]);
    let body=w('tblPr',{},properties(book,props[0],where,indent)+(oldFirst?change(props[0],'tblPr',properties(book,oldFirst,where,oldFirst.edges[0]+oldFirst.padding.left)):''))+w('tblGrid',{},grid.slice(1).map((x,i)=>w('gridCol',{w:x-grid[i]})).join('')),above=new Map(),headers=true;
    for(let i=0;i<rows.length;i++){
      const row=rows[i],p=props[i],place=where+', row '+(i+1),before=indices.get(p.edges[0]),after=grid.length-1-indices.get(p.edges.at(-1)),below=new Map();
      headers=headers&&!!p.header;p.header=headers;
      let rp=rowProperties(book,p,place,before,after,grid)+D.markup.marks(book,row.mark.markProps);
      if(p.revisionFormatting){
        if(p.revisionBefore){rp+=w('trPrChange',D.markup.revision(book,p,'format'),w('trPr',{},rowProperties(book,p.revisionBefore,place,before,after,grid)));book.markupStats.formatRevisions++;}
        else D.notice(book,'legacy-word-revision-format',N('An earlier table formatting change was not kept; the current formatting was retained'),place,true);
      }
      const old=previous(p),sameGrid=old&&old.edges.length===p.edges.length&&old.edges.every((x,j)=>x===p.edges[j]);
      if(old&&!sameGrid)D.notice(book,'legacy-word-revision-grid',N('An earlier table column layout was not kept; the current layout was retained'),place,true);
      if(p.keepFollow)D.notice(book,'legacy-word-table-keep',N('The setting that keeps a table row with the following row was not kept'),place,true);
      let content=(i?w('tblPrEx',{},properties(book,p,place,indent,true)+(old?change(p,'tblPrEx',properties(book,old,place,old.edges[0]+old.padding.left,true)):'')):'')+(rp?w('trPr',{},rp):'');
      for(let j=0;j<row.cells.length;j++){
        const c=p.cells[j],a=indices.get(p.edges[j]),z=indices.get(p.edges[j+1]),key=a+':'+z;
        let vertical=c.vmerge===3?'restart':c.vmerge===1?'continue':null;
        if(vertical==='continue'&&!above.has(key)){vertical=null;D.notice(book,'legacy-word-table-merge',N('A damaged cell merge was removed; its text was retained'),place,true);}
        if(vertical)below.set(key,true);
        let data=blocks(row.cells[j]);data=tail(data,j+1===row.cells.length?row.mark.inner:'');
        let cp=cellProperties(book,c,place,z-a,p.edges[j+1]-p.edges[j],vertical);
        if(sameGrid&&old.cells[j]){const was=old.cells[j];cp+=change(p,'tcPr',cellProperties(book,was,place,z-a,old.edges[j+1]-old.edges[j],was.vmerge===3?'restart':was.vmerge===1?'continue':null));}
        else if(old&&!sameGrid&&old.cells[j]&&JSON.stringify(old.cells[j])!==JSON.stringify(c))D.notice(book,'legacy-word-revision-cells',N('Earlier cell formatting changes were not kept; the current formatting was retained'),place,true);
        content+=w('tc',{},w('tcPr',{},cp)+data);book.tableStats.cells++;
      }
      above=below;body+=w('tr',p.rsid==null?{}:{rsidRPr:p.rsid.toString(16).padStart(8,'0').toUpperCase()},content);book.tableStats.rows++;
    }
    book.tableStats.tables++;return w('tbl',{},body);
  }
  T.story=function(book,nodes,where){
    book.tableStats??={tables:0,rows:0,cells:0,tableFallbacks:0};
    const out=[],stack=[];let count=0;
    const row=()=>({cells:[],cell:[],mark:null});
    function close(){const t=stack.pop();if(t.row.cell.length)t.row.cells.push(t.row.cell);if(t.row.cells.length)t.rows.push(t.row);}
    for(const n of nodes){
      let depth=T.depth(n.p);if(depth>B.limits.depth){D.notice(book,'legacy-word-tables',N('A table nested too deeply was replaced with its cell text'),where,true);depth=0;}
      while(stack.length>depth)close();
      while(stack.length<depth){
        const t={kind:'table',depth:stack.length+1,rows:[],row:row(),where:where+', table '+(++count)};
        (stack.length?stack.at(-1).row.cell:out).push(t);stack.push(t);
      }
      if(!depth){out.push(n);if(n.mark==='\x07')D.notice(book,'legacy-word-tables',N('Tables were replaced with their cell text'),where,true);continue;}
      const t=stack.at(-1);
      if(T.rowEnd(n.p,n.mark)){
        // A valid row terminator is an otherwise empty paragraph. Unexpected
        // text stays in the local fallback instead of being discarded.
        if(n.z-n.a>1)t.row.damaged=true;
        if(t.row.cell.length)t.row.cells.push(t.row.cell);t.row.mark=n;t.rows.push(t.row);t.row=row();
      }else{t.row.cell.push(n);if(cellEnd(n)){t.row.cells.push(t.row.cell);t.row.cell=[];}}
    }
    while(stack.length)close();
    return out.map(n=>n.kind==='table'?write(book,n):n.xml).join('');
  };
})(typeof window!=='undefined'?window:globalThis);
