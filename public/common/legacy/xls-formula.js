/* Derived from Apache OpenOffice main/sc/source/filter/excel/excform8.cxx (Apache-2.0).
 * Derived from Apache POI ss/formula/{FormulaRenderer,SharedFormula}.java, ptg/ and
 * constant/ConstantValueParser.java (Apache-2.0). BIFF RPN -> Excel formula text.
 * This parser never evaluates a formula or follows an external reference. */
(function(root){
  'use strict';
  const B=root.L.legacy,X=B.xls;
  const minimumArgs=new Map(Object.values(X.functions).map(f=>[f[0],f[1]]));
  // Excel: reference operators bind above signs, percent and exponentiation.
  // Binary operators (including ^) associate left; tParen retains typed grouping.
  const precedence={'<':1,'<=':1,'=':1,'>=':1,'>':1,'<>':1,'&':2,'+':3,'-':3,'*':4,'/':4,'^':5,',':8,' ':9,':':10};
  X.formulaErrors={0:'#NULL!',7:'#DIV/0!',15:'#VALUE!',23:'#REF!',29:'#NAME?',36:'#NUM!',42:'#N/A',43:'#GETTING_DATA'};
  // SheetNameFormatter (POI): delimit only ambiguous names. The destination is
  // OOXML, so A1-looking names use the modern sheet bounds, not BIFF's bounds.
  X.quoteSheet=function(name){
    name=String(name);const bare=name.replace(/^\[\d+\]/,'').split(':');
    const ambiguous=s=>{
      if(!s||/^\p{N}/u.test(s)||/[^\p{L}\p{N}_.]/u.test(s)||/^(?:TRUE|FALSE|R|C|RC)$/i.test(s)||/^R(?:\d|C\d)/i.test(s))return true;
      const m=/^([A-Z]{1,3})([1-9]\d*)$/i.exec(s);if(!m)return false;
      const col=[...m[1].toUpperCase()].reduce((n,c)=>n*26+c.charCodeAt(0)-64,0);return col<=16384&&+m[2]<=1048576;
    };
    return bare.some(ambiguous)?"'"+name.replace(/'/g,"''")+"'":name;
  };
  const colName=col=>{let out='';for(let n=col+1;n;n=Math.floor((n-1)/26))out=String.fromCharCode(65+(n-1)%26)+out;return out;};
  const string=s=>'"'+s.replace(/"/g,'""')+'"';
  const bad=(message,where)=>B.fail('FORMULA_UNSUPPORTED',message,where);
  // Used by array constants and external-workbook cached cells. Cursor is local.
  X.formulaConstant=function(r,cursor,b8=true,cp=1252){
    const type=r.u8(cursor.p++),p=cursor.p;let value;
    if(type===2){const item=X.string(r,p,b8,cp);cursor.p=item.end;return {type:'str',value:item.value};}
    r.check(p,8);cursor.p+=8;
    if(type===0)return {type:'blank',value:null};
    if(type===1){value=r.f64(p);if(!Number.isFinite(value))bad('Invalid formula constant');return {type:'n',value};}
    if(type===4)return {type:'b',value:!!r.u8(p)};
    if(type===16){value=X.formulaErrors[r.u8(p)];if(!value)bad('Unknown formula error constant');return {type:'e',value};}
    bad('Unknown formula constant type',type);
  };
  X.formula=function(tokens,extra=new Uint8Array(),context={}){
    const r=new B.Reader(tokens,'Formula tokens'),tail=new B.Reader(extra,'Formula extra data'),b8=context.b8!==false,cp=context.cp||1252;
    const stack=[],extensions=[],originRow=context.row||0,originCol=context.col||0;let p=0,nodes=0;
    const atom=(text,prec=11)=>({text,prec}),push=node=>{if(++nodes>16384||stack.length>=1024)bad('Formula operand limit');stack.push(node);};
    const take=count=>{if(count<0||stack.length<count)bad('Formula operand stack underflow',p);return stack.splice(stack.length-count,count);};
    const readRef=(row,col,relative)=>{
      const rowRel=!!((b8?col:row)&0x8000),colRel=!!((b8?col:row)&0x4000),rowMask=b8?65535:16383;
      row&=rowMask;col&=255;
      if(relative){if(rowRel)row=(row+originRow)&rowMask;if(colRel)col=(col+originCol)&255;}
      return {row,col,rowRel,colRel,text:(colRel?'':'$')+colName(col)+(rowRel?'':'$')+(row+1)};
    };
    const reference=(area,relative)=>{
      const row1=r.u16(p);p+=2;let a,b;
      if(area){const row2=r.u16(p);p+=2;const col1=b8?r.u16(p):r.u8(p);p+=b8?2:1;const col2=b8?r.u16(p):r.u8(p);p+=b8?2:1;a=readRef(row1,col1,relative);b=readRef(row2,col2,relative);}
      else{const col=b8?r.u16(p):r.u8(p);p+=b8?2:1;a=readRef(row1,col,relative);}
      if(!b)return a.text;
      if(!a.colRel&&!b.colRel&&a.col===0&&b.col===255)return (a.rowRel?'':'$')+(a.row+1)+':'+(b.rowRel?'':'$')+(b.row+1);
      if(!a.rowRel&&!b.rowRel&&a.row===0&&b.row===(b8?65535:16383))return (a.colRel?'':'$')+colName(a.col)+':'+(b.colRel?'':'$')+colName(b.col);
      return a.text+':'+b.text;
    };
    const call=(id,count)=>{
      const fn=X.functions[id];if(id!==255&&!fn)bad('Unknown Excel function',id);
      const args=take(count==null?fn[1]:count);let name=fn?.[0];
      if(id===255){const first=args.shift();name=first?.name||first?.text;if(!name||!/^(?:(?:'(?:[^']|'')*'|\[\d+\](?:[\p{L}\p{N}_.]+)?|[\p{L}_][\p{L}\p{N}_.]*)!)?[_\p{L}\\][\p{L}\p{N}_.\\]*$/u.test(name))bad('Invalid named function');
        // An unbound registered UDF named e.g. Sum is not BIFF's SUM function.
        // Plain A1 text would silently rebind it to the built-in implementation.
        if(first.functionOnly&&minimumArgs.has(name.toUpperCase()))bad('A registered function has the same name as an Excel function',name);
      }
      // Excel rejects known functions with missing required operands, including
      // a NAME registration colliding with a built-in (for example SUMIF(x)).
      const minimum=minimumArgs.get(name.toUpperCase().replace(/^_XLFN\.|^_XLL\./,''));
      if(minimum!=null&&args.length<minimum)bad('A function has too few arguments',name);
      // The Office 2016 plain-workbook gate refuses this Excel 4 macro name.
      // Keep its dependent cells' caches until macro-capable conversion exists.
      if(context.nameMode&&name.toUpperCase()==='GET.CELL')bad('Excel macro-function names cannot be converted yet',name);
      push({fn:name,args});
    };
    const binary={3:'+',4:'-',5:'*',6:'/',7:'^',8:'&',9:'<',10:'<=',11:'=',12:'>=',13:'>',14:'<>',15:' ',16:',',17:':'};
    while(p<r.length){
      const raw=r.u8(p++),id=raw>=0x20?(raw&31)|0x20:raw,relative=!!(context.shared||context.nameMode);
      if(binary[id]){push({op:binary[id],args:take(2)});continue;}
      if(id===0x12||id===0x13||id===0x14||id===0x15){push({unary:id,args:take(1)});continue;}
      if(id===0x16){push(atom(''));continue;}
      if(id===0x17){const item=X.string(r,p,b8,cp,true);p=item.end;push(atom(string(item.value)));continue;}
      if(id===0x19){const flags=r.u8(p),data=r.u16(p+1);p+=3;if(flags&4)p+=(data+1)*2;if(flags&16)push({fn:'SUM',args:take(1)});r.check(p,0);continue;}
      if(id===0x1c){const error=X.formulaErrors[r.u8(p++)];if(!error)bad('Unknown formula error');push(atom(error));continue;}
      if(id===0x1d){push(atom(r.u8(p++)?'TRUE':'FALSE'));continue;}
      if(id===0x1e){push(atom(String(r.u16(p))));p+=2;continue;}
      if(id===0x1f){const value=r.f64(p);p+=8;if(!Number.isFinite(value))bad('Invalid formula number');push(atom(String(value),value<0?7:11));continue;}
      if(id===0x20){r.check(p,7);p+=7;const node=atom('');extensions.push({kind:'array',node});push(node);continue;}
      if(id===0x21){const fn=r.u16(p);p+=2;call(fn);continue;}
      if(id===0x22){const count=r.u8(p)&127,fn=r.u16(p+1);p+=3;call(fn,count);continue;}
      if(id===0x23){const index=r.u16(p);p+=b8?4:14;r.check(p,0);const name=context.name?.(index);if(!name)bad('Undefined formula name',index);push(typeof name==='string'?{text:name,name}:name);continue;}
      if([0x24,0x25,0x2c,0x2d].includes(id)){
        const area=id===0x25||id===0x2d,ref=reference(area,relative||id===0x2c||id===0x2d);
        // MS-OE376 3.2.3.1.4: a name formula's current-sheet reference is
        // !A1, not bare A1. Its sheet remains relative to the caller.
        push(atom((context.nameMode?'!':'')+ref,area?10:11));continue;
      }
      if([0x26,0x27,0x28].includes(id)){r.check(p,6);p+=6;if(id===0x26)extensions.push({kind:'memory'});continue;}
      if(id===0x29||id===0x2e||id===0x2f){r.check(p,2);p+=2;continue;}
      if(id===0x2a||id===0x2b){const n=(id===0x2a?1:2)*(b8?4:3);r.check(p,n);p+=n;push(atom('#REF!'));continue;}
      if(id===0x39){if(!b8)bad('BIFF5 external-name token');const sheet=r.u16(p),index=r.u16(p+2);p+=6;r.check(p,0);const name=context.externalName?.(sheet,index);if(!name)bad('Undefined external formula name',index);push(typeof name==='string'?{text:name,name}:name);continue;}
      if([0x3a,0x3b,0x3c,0x3d].includes(id)){
        if(!b8)bad('BIFF5 external-sheet token');const index=r.u16(p);p+=2;
        // ExtPtgRef3D/ExtPtgArea3D in an external name contain two direct
        // SUPBOOK sheet indices, not the workbook's one-index XTI reference.
        let last=index;if(context.externalDefinition){last=r.u16(p);p+=2;}
        const area=id===0x3b||id===0x3d,ref=reference(area,relative);
        if(id===0x3c||id===0x3d){push(atom('#REF!'));continue;}
        const prefix=context.sheet?.(index,last);if(prefix==null)bad('Undefined external sheet reference',index);push(atom(prefix==='#REF!'?'#REF!':prefix+'!'+ref,area&&prefix!=='#REF!'?10:11));continue;
      }
      bad('Unsupported formula token',raw);
    }
    const cursor={p:0};
    for(const extension of extensions){
      if(extension.kind==='memory'){const n=tail.u16(cursor.p);cursor.p+=2;tail.check(cursor.p,n*8);cursor.p+=n*8;continue;}
      const cols=tail.u8(cursor.p)+1,rows=tail.u16(cursor.p+1)+1;cursor.p+=3;
      if(rows*cols>4096)bad('Formula array exceeds text limit');const items=[];let length=0;
      for(let row=0;row<rows;row++){
        const values=[];for(let col=0;col<cols;col++){
          const c=X.formulaConstant(tail,cursor,b8,cp),value=c.type==='str'?string(c.value):c.type==='blank'?'':c.type==='b'?c.value?'TRUE':'FALSE':String(c.value);
          if((length+=value.length+1)>8192)bad('Formula array exceeds text limit');values.push(value);
        }items.push(values.join(','));
      }extension.node.text='{'+items.join(';')+'}';
    }
    if(stack.length!==1)bad('Unbalanced formula operands');
    const operand=(node,min=0,union=true)=>({node,min,union});
    const parts=[],work=[operand(stack[0])];let length=0;
    while(work.length){
      const item=work.pop();if(typeof item==='string'){length+=item.length;if(length>8192)bad('Formula exceeds Excel text limit');parts.push(item);continue;}
      const {node,min,union}=item,prec=node.op?precedence[node.op]:node.unary===0x14?6:node.unary&&node.unary!==0x15?7:node.prec??11;
      // A bare union inside an argument would become an argument separator.
      if(prec<min||node.op===','&&!union){work.push(')',operand(node),'(');continue;}
      if(node.functionOnly)bad('A formula refers to an undefined name',node.name);
      if(node.text!=null){work.push(node.text);continue;}
      if(node.op){work.push(operand(node.args[1],prec+1,union),node.op,operand(node.args[0],prec,union));continue;}
      if(node.unary){const id=node.unary;
        if(id===0x15)work.push(')',operand(node.args[0]),'(');
        else if(id===0x14)work.push('%',operand(node.args[0],prec,union));
        else work.push(operand(node.args[0],prec,union),id===0x12?'+':'-');
        continue;
      }
      work.push(')');for(let i=node.args.length-1;i>=0;i--){work.push(operand(node.args[i],0,false));if(i)work.push(',');}work.push(node.fn+'(');
    }
    const text=parts.join('');if(/[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff]/.test(text))bad('Formula contains invalid XML characters');return text;
  };
})(typeof window!=='undefined'?window:globalThis);
