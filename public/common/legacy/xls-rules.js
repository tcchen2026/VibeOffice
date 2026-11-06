/* Derived from Apache POI hssf/record/{DV,CFHeader,CFRule,CFRule12}Record.java,
 * cf/{Threshold,ColorGradientFormatting,DataBarFormatting,IconMultiStateFormatting}.java
 * (Apache-2.0). Derived from Apache OpenOffice main/sc/source/filter/excel/xicontent.cxx
 * (Apache-2.0). MS-XLS CFEx/CFExNonCF12 supply priority, activity and format overrides. */
(function(root){
  'use strict';
  const B=root.L.legacy,X=B.xls,el=B.el,N=s=>s,bad=s=>B.fail('XLS_RULE',s);
  const operators=['between','notBetween','equal','notEqual','greaterThan','lessThan','greaterThanOrEqual','lessThanOrEqual'];
  const rangeList=(r,p)=>{const n=r.u16(p);p+=2;r.check(p,n*8);if(!n)bad('A rule has no cell range');return Array.from({length:n},(_,i)=>X.cellRange(r,p+i*8));};
  const formula=(bytes,ctx)=>X.formula(bytes,undefined,{...ctx,shared:true});
  X.validation=function(r,context){
    const flags=r.u32(0),type=['none','whole','decimal','list','date','time','textLength','custom'][flags&15];if(!type)bad('Unknown data validation type');
    const attrs={type,operator:type==='list'||type==='custom'||type==='none'?null:operators[flags>>>20&7],errorStyle:['stop','warning','information'][flags>>>4&7],allowBlank:flags&256?1:0,showDropDown:flags&512?1:null,showInputMessage:flags&0x40000?1:0,showErrorMessage:flags&0x80000?1:0};let p=4;
    for(const tag of ['promptTitle','errorTitle','prompt','error']){const s=X.string(r,p);p=s.end;if(s.value!=='\0'&&s.value)attrs[tag]=s.value;}
    const tokens=[];for(let i=0;i<2;i++){const n=r.u16(p);p+=4;tokens.push(r.sub(p,n));p+=n;}
    const ranges=rangeList(r,p),[row,col]=ranges[0],ctx=context(row,col);attrs.sqref=ranges.map(a=>X.range(...a)).join(' ');
    const formulas=tokens.map((t,i)=>{
      if(!t.length)return '';
      if(i===0&&type==='list'&&flags&128){const value=new B.Reader(t);if(value.u8(0)!==0x17)bad('Invalid explicit validation list');return '"'+X.string(value,1,true,bookCodepage(ctx),true).value.replace(/\0/g,',').replace(/"/g,'""')+'"';}
      return formula(t,ctx);
    });
    if(type!=='none'&&!formulas[0])bad('The validation condition is missing');
    const body=formulas.map((f,i)=>f?el('formula'+(i+1),{},B.xstring(f)):'').join('');return el('dataValidation',attrs,body);
  };
  const bookCodepage=ctx=>ctx.cp||1252;
  const icons=['3Arrows','3ArrowsGray','3Flags','3TrafficLights1','3Signs','3TrafficLights2','3Symbols','3Symbols2','4Arrows','4ArrowsGray','4RedToBlack','4Rating','4TrafficLights','5Arrows','5ArrowsGray','5Rating','5Quarters'];
  const threshold=function(r,cursor,ctx,kind){
    let p=cursor.p;const type=({1:'num',2:'min',3:'max',4:'percent',5:'percentile',7:'formula'})[r.u8(p)],n=r.u16(p+1);p+=3;if(!type)bad('Unknown conditional-format threshold');
    const attrs={type};if(n){attrs.val=formula(r.sub(p,n),ctx);p+=n;}else if(type!=='min'&&type!=='max'){const v=r.f64(p);p+=8;if(!Number.isFinite(v))bad('Invalid conditional-format threshold');attrs.val=v;}
    if(kind==='color'){r.check(p,8);p+=8;}else if(kind==='icon'){attrs.gte=r.u8(p)?null:0;r.check(p,5);p+=5;}
    cursor.p=p;return el('cfvo',attrs);
  };
  const modern=function(r,p,type,ctx,book){
    const cursor={p:p+6};r.check(p,6);
    if(type===3){
      const n=r.u8(p+3),colors=r.u8(p+4);if(n<2||n>3||n!==colors)bad('Unsupported color-scale size');const stops=Array.from({length:n},()=>threshold(r,cursor,ctx,'color')),values=[];
      for(let i=0;i<n;i++){cursor.p+=8;values.push(el('color',X.cfColor(r,cursor.p,book)));cursor.p+=16;}
      return {type:'colorScale',body:el('colorScale',{},stops.join('')+values.join(''))};
    }
    if(type===4){
      const flags=r.u8(p+3),minLength=r.u8(p+4),maxLength=r.u8(p+5);if(minLength>maxLength||maxLength>100)bad('Invalid data-bar lengths');
      const color=el('color',X.cfColor(r,cursor.p,book));cursor.p+=16;const stops=threshold(r,cursor,ctx)+threshold(r,cursor,ctx);
      // BIFF's right-to-left flag needs the x14 data-bar extension. Fail locally
      // instead of silently reversing the visual comparison.
      if(flags&1)bad('Right-to-left data bars cannot be converted yet');
      return {type:'dataBar',body:el('dataBar',{minLength,maxLength,showValue:flags&2?null:0},stops+color)};
    }
    if(type===6){
      const count=r.u8(p+3),name=icons[r.u8(p+4)],flags=r.u8(p+5);if(!name||+name[0]!==count)bad('Unknown conditional icon set');
      return {type:'iconSet',body:el('iconSet',{iconSet:name,showValue:flags&1?0:null,reverse:flags&4?1:null},Array.from({length:count},()=>threshold(r,cursor,ctx,'icon')).join(''))};
    }
    if(type===5){const flags=r.u8(p+3),rank=r.u16(p+4);if(!rank||rank>(flags&2?100:1000))bad('Invalid Top 10 rule');return {type:'top10',bottom:flags&1?null:1,percent:flags&2?1:null,rank,body:''};}
    bad('Unknown conditional formatting type');
  };
  X.sheetRules=function(book,si,meta,formulas,styles,losses){
    const sheet=book.sheets[si],groups=[],byId=new Map(),validations=[],dvAttrs={},stats={validationSource:0,validations:0,cfSource:0,conditionalFormats:0};let current=null,sequence=0;
    const issue=(kind,e,where=sheet.name)=>X.featureLoss(losses,'legacy-'+kind,kind==='validation'?N('A data validation rule was not kept'):kind==='conditional-formatting-settings'?N('Some conditional formatting settings were not kept'):N('A conditional formatting rule was not kept'),where,e);
    const context=(row,col)=>formulas.context(si,row,col);
    for(let i=0;i<meta.records.length;i++){
      const rec=meta.records[i],id=rec.id;if(![0x1b2,0x1be,0x1b0,0x879,0x1b1,0x87a,0x87b].includes(id))continue;
      if(id===0x1be)stats.validationSource++;if(id===0x1b1||id===0x87a)stats.cfSource++;
      try{
        const r=X.continuedRecord(meta.records,i);
        if(id===0x1b2){dvAttrs.disablePrompts=r.u16(0)&1?1:0;dvAttrs.xWindow=r.u32(2);dvAttrs.yWindow=r.u32(6);}
        else if(id===0x1be)validations.push(X.validation(r,context));
        else if(id===0x1b0||id===0x879){
          current=null;const p=id===0x879?12:0,ranges=rangeList(r,p+12),group={id:r.u16(p+2)>>>1,ranges,rules:[]};groups.push(group);byId.set(group.id,group);current=group;
        }else if(id===0x87b){
          const group=byId.get(r.u16(16));if(r.u32(12)){current=group;continue;}
          const rule=group?.rules[r.u16(18)];if(!rule)bad('Missing conditional-format rule for extension');
          const flags=r.u8(24);rule.priority=r.u16(22);rule.stopIfTrue=flags&2?1:null;rule.active=!!(flags&1);
          if(r.u8(25)){const n=r.u32(26);r.check(30,n||2);if(n)rule.dxf=X.dxf(r,30,book,styles,30+n).body;}
        }else{
          if(!current)bad('Missing conditional-format cell range');
          const rule={sequence:sequence++,active:true,stopIfTrue:1},p0=id===0x87a?12:0,type=r.u8(p0),op=r.u8(p0+1),n1=r.u16(p0+2),n2=r.u16(p0+4),[row,col]=current.ranges[0],ctx=context(row,col);let p=p0+6;
          // Keep a slot even when a classic rule is damaged: CFEx indexes it.
          current.rules.push(rule);
          try{
            if(id===0x87a){const size=r.u32(p);p+=4;r.check(p,size||2);if(size)rule.dxf=X.dxf(r,p,book,styles,p+size).body;p+=size||2;}
            else{const dxf=X.dxf(r,p,book,styles);rule.dxf=dxf.body;p=dxf.end;}
            const f1=n1?formula(r.sub(p,n1),ctx):'';p+=n1;const f2=n2?formula(r.sub(p,n2),ctx):'';p+=n2;
            if(id===0x87a){const n=r.u16(p);p+=2;const active=n?formula(r.sub(p,n),ctx):'';p+=n;if(active&&active!=='TRUE'&&active!=='1')bad('A conditional visual rule has an unsupported activity formula');
              rule.stopIfTrue=r.u8(p)&2?1:null;rule.priority=r.u16(p+1);const size=r.u8(p+5);r.check(p+6,size);p+=6+size;
            }
            if(type===1||type===2){rule.type=type===1?'cellIs':'expression';rule.operator=type===1?operators[op-1]:null;if(!f1||type===1&&!rule.operator||type===1&&op<=2&&!f2)bad('A conditional-format condition is missing');rule.body=el('formula',{},B.xml(f1))+(f2?el('formula',{},B.xml(f2)):'');}
            else Object.assign(rule,modern(r,p,type,ctx,book));
          }catch(e){rule.invalid=true;throw e;}
        }
      }catch(e){if(e.name!=='LegacyError')throw e;issue(id===0x1be||id===0x1b2?'validation':id===0x87b?'conditional-formatting-settings':'conditional-formatting',e);}
    }
    // Keep relative priority; old BIFF rules have implicit first-to-last order.
    const rules=groups.flatMap(g=>g.rules).filter(r=>!r.invalid&&r.active).sort((a,b)=>(a.priority??a.sequence+1)-(b.priority??b.sequence+1)||a.sequence-b.sequence);
    rules.forEach((r,i)=>r.priority=i+1);stats.conditionalFormats=rules.length;stats.validations=validations.length;
    const conditional=groups.map(g=>{const body=g.rules.filter(r=>!r.invalid&&r.active).map(rule=>{const {sequence,active,invalid,dxf,body,...attrs}=rule;if(dxf)attrs.dxfId=styles.dxf(dxf);return el('cfRule',attrs,body);}).join('');return body?el('conditionalFormatting',{sqref:g.ranges.map(a=>X.range(...a)).join(' ')},body):'';}).join('');
    return {conditional,validation:validations.length?el('dataValidations',{...dvAttrs,count:validations.length},validations.join('')):'',stats};
  };
})(typeof window!=='undefined'?window:globalThis);
