/* Grouped PowerPoint table cells and border lines -> DrawingML tables.
 * Derived from Apache POI hslf/usermodel/{HSLFTable,HSLFTableCell}.java (Apache-2.0).
 * Derived from Apache OpenOffice filter/source/msfilter/svdfppt.cxx CreateTable
 * (Apache-2.0). Grid edges and spans come from the binary anchors. */
(function(root){
  'use strict';
  const B=root.L.legacy,P=B.ppt,D=P.drawings,S=B.shapes,T=P.text,el=B.el,N=s=>s,Q=P.tables={};
  const get=(s,id,fallback=0)=>s.properties.get(id)?.value??fallback;
  const bad=()=>B.fail('PPT_TABLE',N('The table cell grid could not be kept'));
  Q.read=function(shape){
    const cells=[],lines=[];for(const s of shape.children||[]){if(!s.anchor||s.children?.length)bad();if(s.type===20)lines.push(s);else if(s.type===1||s.type===202)cells.push(s);else bad();}
    if(!cells.length)bad();
    const xs=[...new Set(cells.flatMap(s=>[s.anchor.left,s.anchor.right]))].sort((a,b)=>a-b),ys=[...new Set(cells.flatMap(s=>[s.anchor.top,s.anchor.bottom]))].sort((a,b)=>a-b),cols=xs.length-1,rows=ys.length-1;
    if(cols<1||rows<1||cols>1024||rows>10000||cols*rows>100000)bad();
    const xi=new Map(xs.map((v,i)=>[v,i])),yi=new Map(ys.map((v,i)=>[v,i])),grid=new Int32Array(cols*rows).fill(-1),items=[];
    for(const s of cells){const a=s.anchor,c=xi.get(a.left),r=yi.get(a.top),cx=xi.get(a.right),ry=yi.get(a.bottom);if(cx<=c||ry<=r)bad();const item={shape:s,r,c,rows:ry-r,cols:cx-c};
      for(let y=r;y<ry;y++)for(let x=c;x<cx;x++){const i=y*cols+x;if(grid[i]!==-1)bad();grid[i]=items.length;}items.push(item);
    }
    if(grid.some(i=>i<0))bad();
    const edges=new Map();let budget=200000;
    const edge=(r,c,side,line)=>{if(r<0||r>=rows||c<0||c>=cols)return;if(--budget<0)bad();edges.set(r+':'+c+':'+side,line);};
    for(const line of lines){const a=line.anchor,x1=Math.min(a.left,a.right),x2=Math.max(a.left,a.right),y1=Math.min(a.top,a.bottom),y2=Math.max(a.top,a.bottom);
      if(x1===x2&&xi.has(x1)&&yi.has(y1)&&yi.has(y2)){const c=xi.get(x1);for(let r=yi.get(y1);r<yi.get(y2);r++){edge(r,c-1,'R',line);edge(r,c,'L',line);}}
      else if(y1===y2&&yi.has(y1)&&xi.has(x1)&&xi.has(x2)){const r=yi.get(y1);for(let c=xi.get(x1);c<xi.get(x2);c++){edge(r-1,c,'B',line);edge(r,c,'T',line);}}
      else bad();
    }
    return {xs,ys,rows,cols,grid,items,edges};
  };
  Q.prepare=function(book){
    const walk=(shapes,page)=>{for(const s of shapes){if(get(s,927)&1)try{s.tableData=Q.read(s);}catch(e){if(e.name!=='LegacyError')throw e;D.report(book,page,'table',s,e);}walk(s.children||[],page);}};
    for(const page of D.pages(book))walk(page.textShapes||[],page);
  };
  Q.write=function(book,page,shape,rect,id,cNvPr){
    const q=shape.tableData;if(!q)return null;
    // A graphicFrame has no independent rotation. Keep a rotated source table
    // as its original group until it can be wrapped without changing IDs.
    if(get(shape,4)||shape.flags&192){D.report(book,page,'table',shape,'Rotated table');return null;}
    const {xs,ys,cols,rows,grid,items,edges}=q,sx=rect.cx/(xs.at(-1)-xs[0]),sy=rect.cy/(ys.at(-1)-ys[0]);
    if(!(sx>0&&sy>0))bad();
    const context={scheme:page.scheme,imageRel:book.drawingImageRel,components:true},empty=el('a:txBody',{},el('a:bodyPr')+el('a:lstStyle')+el('a:p',{},el('a:endParaRPr')));
    const body=[];
    for(let r=0;r<rows;r++){
      let cells='';for(let c=0;c<cols;c++){
        const item=items[grid[r*cols+c]],s=item.shape,origin=item.r===r&&item.c===c,attrs=origin?{rowSpan:item.rows>1?item.rows:null,gridSpan:item.cols>1?item.cols:null}:{hMerge:c>item.c?1:null,vMerge:r>item.r?1:null};
        const margins={marL:get(s,129,91440)|0,marR:get(s,131,91440)|0,marT:get(s,130,45720)|0,marB:get(s,132,45720)|0,anchor:['t','ctr','b','t','ctr','b'][get(s,135)]||'t',anchorCtr:get(s,135)>=3?1:null,vert:['horz','vert','vert270','vert','horz','vert'][get(s,136)]||'horz'};
        let style;try{style=S.style(s,context);}catch(e){if(e.name!=='LegacyError')throw e;D.report(book,page,'format',s,e);style={fill:el('a:noFill'),effects:''};}
        let borders='';for(const side of ['L','R','T','B']){const border=edges.get(r+':'+c+':'+side)||s;try{borders+=S.style(border,{...context,lineTag:'a:ln'+side}).line;}catch(e){if(e.name!=='LegacyError')throw e;D.report(book,page,'format',border,e);borders+=el('a:ln'+side,{},el('a:noFill'));}}
        if(style.effects)D.report(book,page,'format',s,'Table cell shadow');
        cells+=el('a:tc',attrs,(origin?T.body(s,book,page,false,{tag:'a:txBody',bodyPr:el('a:bodyPr')}):empty)+el('a:tcPr',margins,borders+style.fill));
      }
      body.push(el('a:tr',{h:Math.max(1,Math.round((ys[r+1]-ys[r])*sy))},cells));
    }
    for(const s of shape.children||[])s.drawingWritten=true;
    const table=el('a:tbl',{},el('a:tblPr')+el('a:tblGrid',{},xs.slice(1).map((x,i)=>el('a:gridCol',{w:Math.max(1,Math.round((x-xs[i])*sx))})).join(''))+body.join(''));
    return el('p:graphicFrame',{},el('p:nvGraphicFramePr',{},cNvPr+el('p:cNvGraphicFramePr',{},el('a:graphicFrameLocks',{noGrp:1}))+el('p:nvPr'))+el('p:xfrm',{},el('a:off',{x:Math.round(rect.x),y:Math.round(rect.y)})+el('a:ext',{cx:Math.round(rect.cx),cy:Math.round(rect.cy)}))+el('a:graphic',{},el('a:graphicData',{uri:'http://schemas.openxmlformats.org/drawingml/2006/table'},table)));
  };
})(typeof window!=='undefined'?window:globalThis);
