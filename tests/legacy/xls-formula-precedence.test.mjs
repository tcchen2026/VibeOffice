import {test} from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import {load} from '../../tools/legacy/load.mjs';
import {words} from './xls-fixtures.mjs';
const X=load().xls;
if(process.env.XLS_PARENS_BEFORE)vm.runInThisContext(fs.readFileSync(process.env.XLS_PARENS_BEFORE,'utf8'));
const cat=(...parts)=>Buffer.concat(parts.map(p=>Buffer.from(p)));
const n=value=>cat([0x1e],words(value)),cell=col=>cat([0x24],words(0,0xc000|col));
const render=(...parts)=>X.formula(cat(...parts));

test('binary formula text follows Excel precedence and left associativity without outer parentheses',()=>{
  assert.equal(render(n(2),n(3),n(4),[5,3]),'2+3*4');
  assert.equal(render(n(2),n(3),[3],n(4),[5]),'(2+3)*4');
  assert.equal(render(n(2),n(3),[4],n(4),[4]),'2-3-4');
  assert.equal(render(n(2),n(3),n(4),[4,4]),'2-(3-4)');
  assert.equal(render(n(2),n(3),n(4),[3,3]),'2+(3+4)');
  assert.equal(render(n(2),n(3),n(4),[5,6]),'2/(3*4)');
  assert.equal(render(n(2),n(3),[7],n(2),[7]),'2^3^2');
  assert.equal(render(n(2),n(3),n(2),[7,7]),'2^(3^2)');
  assert.equal(render(n(2),n(3),[8],n(4),[11]),'2&3=4');
  assert.equal(render(n(2),n(3),[11],n(4),[8]),'(2=3)&4');
  assert.equal(render(n(2),n(3),n(4),[11,11]),'2=(3=4)');
});

test('typed tParen tokens survive even when redundant and do not gain a second grouping pair',()=>{
  assert.equal(render(n(2),[0x15]),'(2)');
  assert.equal(render(n(2),n(3),[3,0x15,0x15]),'((2+3))');
  assert.equal(render(n(2),n(3),[3,0x15],n(4),[5]),'(2+3)*4');
  assert.equal(render(n(2),n(3),n(4),[4,0x15,4]),'2-(3-4)');
});

test('unary signs and percent keep their own Excel precedence',()=>{
  assert.equal(render(n(2),n(3),[3,0x13]),'-(2+3)');
  assert.equal(render(n(2),[0x13],n(2),[7]),'-2^2');
  assert.equal(render(n(2),n(2),[7,0x13]),'-(2^2)');
  assert.equal(render(n(50),[0x14]),'50%');
  assert.equal(render(n(50),[0x14,0x13]),'-(50%)');
  assert.equal(render(n(50),[0x13,0x14]),'-50%');
  assert.equal(render(n(50),[0x14,0x14]),'50%%');
  assert.equal(render(n(2),n(3),[7,0x14]),'(2^3)%');
  assert.equal(render(n(2),[0x14],n(3),[7]),'2%^3');
  assert.equal(render(n(2),[0x12,0x13]),'-+2');
});

test('reference operators bind above unary and keep comma unions inside one function argument',()=>{
  assert.equal(render(cell(0),cell(1),[0x11,0x13]),'-A1:B1');
  assert.equal(render(cell(0),cell(1),[0x10],cell(2),[0x0f]),'(A1,B1) C1');
  assert.equal(render(cell(0),cell(1),[0x0f],cell(2),[0x11]),'(A1 B1):C1');
  assert.equal(render(cell(0),cell(1),[0x0f],cell(2),cell(3),[0x11,0x10]),'A1 B1,C1:D1');
  const sum=cat([0x22,1],words(4));
  assert.equal(render(cell(0),cell(1),[0x10],sum),'SUM((A1,B1))');
  assert.equal(render(cell(0),cell(1),[0x10,0x15],sum),'SUM((A1,B1))');
  assert.equal(render(cell(0),cell(1),[0x10],n(2),[3],sum),'SUM((A1,B1)+2)');
  const area=cat([0x25],words(0,0,0xc001,0xc002));
  assert.equal(render(cell(0),area,[0x11]),'A1:(B1:C1)');
});

test('15228-style date arithmetic stays readable with its original typed grouping',()=>{
  const named=i=>cat([0x23],words(i,0)),fn=i=>cat([0x21],words(i));
  const tokens=cat(named(1),fn(69),n(12),[5],named(1),fn(68),[3],named(2),fn(69),n(12),[5,4],named(2),fn(68),[4]);
  assert.equal(X.formula(tokens,undefined,{name:i=>i===1?'RunDat':'ValDat'}),'YEAR(RunDat)*12+MONTH(RunDat)-YEAR(ValDat)*12-MONTH(ValDat)');
});
