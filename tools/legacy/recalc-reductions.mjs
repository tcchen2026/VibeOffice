// Small, reviewable Ledger discrepancies. This records findings, not app fixes.
import fs from 'node:fs';
import path from 'node:path';
import {createRequire} from 'node:module';
const require=createRequire(import.meta.url),L=require('../../tests/ledger/load.js'),M=L.model;
const out=path.resolve(process.argv[2]||'tmp/legacy/artifacts/C2-ledger-reductions');fs.mkdirSync(out,{recursive:true});
const cases=[
  {name:'Unary plus text',formula:'+A1',inputs:{A1:'col'},expected:'col',source:'poi__2501c48d2d40__FormulaEvalTestData.xls, EverythingTests L87/M87; reduced to one text cell'},
  {name:'AND missing argument',formula:'AND(,TRUE)',expected:false,source:'poi__75b0f4f3c558__BooleanFunctionsTestCaseData.xls, Tests B19'},
  {name:'COUNTIF error',formula:'COUNTIF(#REF!,"A")',expected:M.err('#REF!'),source:'libreoffice__10ff90742c1d__forum-mso-en4-258437.xls, Sheet1 P1'},
  {name:'DATEVALUE month name',formula:'DATEVALUE("2020-FEB-01")',expected:43862,source:'poi__9699bae4cb57__63819.xls, Sheet1 C4'},
  {name:'CELL locale string',formula:'CELL("adresse",A1)',inputs:{A1:1},expected:'$A$1',source:'poi__fd58d1ee96e2__StringContinueRecords.xls, file load sequence AW4; French info string, locale-sensitive candidate'}
];
const wb=new M.Workbook();wb.calcPr.fullCalcOnLoad=true;
for(const item of cases){const sh=wb.addSheet(item.name);for(const [address,v]of Object.entries(item.inputs||{})){const {r,c}=L.formula.parseCell(address);sh.put(r,c,{v});}sh.put(0,3,{f:item.formula,v:item.expected});sh.put(2,0,{v:item.source});item.sheet=sh;}
const file=path.join(out,'reductions.xlsx');fs.writeFileSync(file,await L.xlsxWrite.bytes(wb,{type:'xlsx'}));L.calc.recalcAll(wb);
const results=cases.map(({sheet,...item})=>({...item,cell:'D1',actual:sheet.get(0,3).v}));fs.writeFileSync(path.join(out,'results.json'),JSON.stringify(results,null,2)+'\n');
fs.writeFileSync(path.join(out,'reopen.json'),JSON.stringify([{file}],null,2)+'\n');console.log(JSON.stringify(results,null,2));
