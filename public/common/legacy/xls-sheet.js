/* Derived from Apache OpenOffice sc/source/filter/excel/{impop,xiview}.cxx and
 * Derived from Apache POI hssf/record/{Row,ColumnInfo,WindowTwo,Pane,Selection}Record.java
 * (Apache-2.0). BIFF worksheet layout and view records. */
(function(root){
  'use strict';
  const B=root.L.legacy,X=B.xls,N=s=>s;
  X.address=(row,col)=>{let s='';for(let n=col+1;n;n=Math.floor((n-1)/26))s=String.fromCharCode(65+(n-1)%26)+s;return s+(row+1);};
  X.range=(r1,c1,r2,c2)=>X.address(r1,c1)+(r1===r2&&c1===c2?'':':'+X.address(r2,c2));
  X.sheetMetadata=function(sheet,losses){
    const out={rows:new Map(),cols:[],merges:[],selections:[],view:{},format:{defaultRowHeight:12.75},outline:{},records:[]};let depth=0;
    for(const rec of sheet.records){const r=rec.r,id=rec.id;if(id===0x809){depth++;continue;}if(id===0xa){depth--;continue;}if(depth!==1)continue;out.records.push(rec);
      try{
        if(id===0x208){const row=r.u16(0),flags=r.u16(12),extra=r.u16(14);out.rows.set(row,{r:row+1,ht:(r.u16(6)&0x7fff)/20,customHeight:flags&64?1:null,hidden:flags&32?1:null,outlineLevel:flags&7||null,collapsed:flags&16?1:null,sourceXF:flags&128?extra&4095:null,thickTop:extra&0x1000?1:null,thickBot:extra&0x2000?1:null,defaultHeight:!!(r.u16(6)&0x8000)});}
        else if(id===0x7d){const first=r.u16(0),last=Math.min(255,r.u16(2)),flags=r.u16(8);if(first<=last)out.cols.push({min:first+1,max:last+1,width:r.u16(4)/256,sourceXF:r.u16(6),customWidth:1,hidden:flags&1?1:null,outlineLevel:flags>>>8&7||null,collapsed:flags&4096?1:null});}
        else if(id===0x225){const flags=r.u16(0);Object.assign(out.format,{defaultRowHeight:(r.u16(2)&0x7fff)/20,customHeight:flags&1?1:null,zeroHeight:flags&2?1:null,thickTop:flags&4?1:null,thickBottom:flags&8?1:null});}
        else if(id===0x55)out.format.baseColWidth=Math.min(255,r.u16(0));
        else if(id===0x99)out.format.defaultColWidth=r.u16(0)/256;
        else if(id===0x81){const flags=r.u16(0);out.dialog=!!(flags&16);out.outline.summaryBelow=flags&64?1:0;out.outline.summaryRight=flags&128?1:0;}
        else if(id===0xe5){const count=r.u16(0);r.check(2,count*8);for(let i=0;i<count;i++){const p=2+i*8,box=[r.u16(p),r.u16(p+4),r.u16(p+2),r.u16(p+6)];if(box[0]<=box[2]&&box[1]<=box[3]&&box[3]<256)out.merges.push(box);}}
        else if(id===0x23e){const flags=r.u16(0);out.flags=flags;Object.assign(out.view,{showFormulas:flags&1?1:null,showGridLines:flags&2?null:0,showRowColHeaders:flags&4?null:0,showZeros:flags&16?null:0,rightToLeft:flags&64?1:null,showOutlineSymbols:flags&128?null:0,tabSelected:flags&512?1:null,view:flags&2048?'pageBreakPreview':null,topLeftCell:X.address(r.u16(2),Math.min(255,r.u16(4)))});if(!(flags&32)){out.view.defaultGridColor=0;out.view.colorId=r.u16(6)&127;}if(r.length>=14){out.view.zoomScaleSheetLayoutView=r.u16(10)||null;out.view.zoomScaleNormal=r.u16(12)||null;}}
        else if(id===0xa0){const denominator=r.u16(2);if(denominator)out.view.zoomScale=Math.round(r.u16(0)*100/denominator);}
        else if(id===0x41)out.pane={xSplit:r.u16(0),ySplit:r.u16(2),topLeftCell:X.address(r.u16(4),Math.min(255,r.u16(6))),activePane:['bottomRight','topRight','bottomLeft','topLeft'][r.u16(8)&3]};
        else if(id===0x1d){const count=r.u16(7);r.check(9,count*6);const ranges=[];for(let i=0;i<count;i++){const p=9+i*6,r1=r.u16(p),r2=r.u16(p+2),c1=r.u8(p+4),c2=r.u8(p+5);if(r1<=r2&&c1<=c2)ranges.push(X.range(r1,c1,r2,c2));}out.selections.push({pane:['bottomRight','topRight','bottomLeft','topLeft'][r.u8(0)&3],activeCell:X.address(r.u16(1),Math.min(255,r.u16(3))),activeCellId:Math.min(Math.max(0,ranges.length-1),r.u16(5)),sqref:ranges.join(' ')||'A1'});}
      }catch(e){if(e.name!=='LegacyError')throw e;losses.push(B.loss('legacy-sheet-layout',N('Some worksheet layout or view settings could not be read'),rec.offset,'drop',sheet.name));}
    }
    for(const key of ['zoomScale','zoomScaleNormal','zoomScaleSheetLayoutView'])if(out.view[key]!=null)out.view[key]=Math.max(10,Math.min(400,out.view[key]));
    if(out.pane)out.pane.state=out.flags&8?out.flags&256?'frozen':'frozenSplit':'split';
    const unique=new Map();for(const m of out.merges)unique.set(m.join(','),m);out.merges=[...unique.values()].sort((a,b)=>a[0]-b[0]||a[1]-b[1]);
    let active=[],work=0;for(const m of out.merges){active=active.filter(a=>a[2]>=m[0]);for(const a of active){if(++work>B.limits.records)B.fail('MERGE_LIMIT',N('There are too many overlapping merged-cell ranges'));if(a[1]<=m[3]&&m[1]<=a[3])B.fail('MERGE_OVERLAP',N('The worksheet contains overlapping merged cells'),sheet.name);}active.push(m);}
    // Keep column defaults sparse. Later records replace earlier ranges, as Excel does.
    const columns=Array(256);for(const c of out.cols)for(let i=c.min-1;i<c.max;i++)columns[i]=c;
    out.cols=[];for(let i=0;i<256;){const c=columns[i],start=i++;while(i<256&&columns[i]===c)i++;if(c)out.cols.push({...c,min:start+1,max:i});}
    return out;
  };
})(typeof window!=='undefined'?window:globalThis);
