/**
 * LAMBDA and the functions that call one: a LAMBDA called where it is written, with fewer or more arguments, left out
 * ones and ISOMITTED, one left without a call, one in LET, one made by another and keeping its LET's names, references
 * as arguments, recursion through LET to find Excel's depth; a function by its name as a value; MAP, REDUCE, SCAN,
 * BYROW, BYCOL and MAKEARRAY over ranges and arrays, with errors inside, answers that are not one value, wrong counts.
 * An array answer is read through ARRAYTOTEXT (strict), so its size and values are one case. Formulas in the engine's
 * spelling, written as a file writes them (`_xlpm.` names, `_xleta.` functions). Answers:
 * `test/excel/cases/lambda.json`; `PREDICTED`, the engine's answers (`sets/predict.ts lambda`), was committed before Excel
 * ran the set.
 */
import { readFileSync } from 'node:fs'
import type { Val } from '../formulajs.js'
import type { Probe } from '../probes.js'
import { xl } from './arrays/common.js'
import { xlpm } from './let/xlpm.js'

/** An array answer as one value. */
const text = (f: string) => `ARRAYTOTEXT(${f},1)`

/** A LAMBDA that counts down from n through LET, `n` calls deep. */
const deep = (n: number) => `LET(f,LAMBDA(g,n,IF(n<=0,0,1+g(g,n-1))),f(f,${n}))`

const CASES: Record<string, string[]> = {
  LAMBDA: [
    'LAMBDA(x,x*2)(3)',
    'LAMBDA(x,y,x+y)(1,2)',
    'LAMBDA(x,x)(1,2)',
    'LAMBDA(x,y,x+y)(1)',
    'LAMBDA(x,y,x+y)(1,)',
    'LAMBDA(x,y,y)(1)',
    'LAMBDA(x,y,ISBLANK(y))(1,)',
    'LAMBDA(x,y,TYPE(y))(1,)',
    'LAMBDA(x,y,y&"")(1,)',
    'LAMBDA(x,x)',
    'LAMBDA(1)',
    'LAMBDA(1)()',
    'LAMBDA(x,x)()',
    'TYPE(LAMBDA(x,x))',
    'ISERROR(LAMBDA(x,x))',
    'ERROR.TYPE(LAMBDA(x,x))',
    'IFERROR(LAMBDA(x,x),"e")',
    'LAMBDA(x,x)+1',
    'LET(f,LAMBDA(x,x+1),f(2))',
    'LET(f,LAMBDA(x,x+1),f)',
    'LAMBDA(x,LAMBDA(y,x+y))(1)(2)',
    'LET(a,10,LAMBDA(x,x+a))(5)',
    'LET(f,LAMBDA(x,x*2),g,LAMBDA(x,f(x)+1),g(3))',
    'LAMBDA(r,ROWS(r))(D!B1:B8)',
    'LAMBDA(r,ISREF(r))(D!B1:B8)',
    'LAMBDA(r,SUM(r))(D!B1:B8)',
    text('LAMBDA(x,x*2)({1,2,3})'),
    'LAMBDA(x,x)(D!A3)',
    'LAMBDA(x,x)(D!A1)',
    'LAMBDA(x,x)(D!A7)',
    'LAMBDA(a1,a1*2)(4)',
    'LAMBDA(_x,_x+1)(1)',
    'LAMBDA(x,X+1)(1)',
    'LAMBDA(sum,sum+1)(1)',
    'LAMBDA(sum,SUM(sum,1))(2)',
    'LAMBDA(x,x)(LAMBDA(y,y*3))(2)',
    'CHOOSE(2,LAMBDA(x,x),LAMBDA(x,x*10))(3)',
    '(LAMBDA(x,x*2))(3)',
    'IF(TRUE,LAMBDA(x,x+1),0)(1)',
    'LAMBDA(x,y,z,x&y&z)("a","b","c")',
    deep(100),
    deep(500),
    deep(1000),
    deep(1023),
    deep(1024),
    deep(2000),
    deep(5000),
  ],
  ISOMITTED: [
    'ISOMITTED(1)',
    'LAMBDA(x,ISOMITTED(x))()',
    'LAMBDA(x,y,ISOMITTED(y))(1,)',
    'LAMBDA(x,y,ISOMITTED(y))(1,2)',
    'LAMBDA(x,y,ISOMITTED(y))(1)',
    'LAMBDA(x,y,ISOMITTED(y))(1,D!A1)',
    'LAMBDA(x,LAMBDA(y,ISOMITTED(y))(x))()',
    'LET(x,1,ISOMITTED(x))',
    'LAMBDA(x,y,IF(ISOMITTED(y),x,x+y))(5)',
    'LAMBDA(x,y,IF(ISOMITTED(y),x,x+y))(5,1)',
  ],
  SUM: ['LET(f,SUM,f(1,2))', 'SUM', 'LET(f,SUM,f())', 'LET(f,SUM,f)'],
  ABS: ['LET(f,ABS,f(-2))'],
  IF: ['LET(f,IF,f(TRUE,1,2))'],
  MAP: [
    text('MAP(D!B1:B4,LAMBDA(x,x*2))'),
    text('MAP(D!B1:C2,LAMBDA(x,x+1))'),
    text('MAP(D!B1:B3,D!C1:C3,LAMBDA(a,b,a*b))'),
    text('MAP(D!B1:B3,D!C1:C2,LAMBDA(a,b,a+b))'),
    text('MAP(D!B1:B2,D!B1:C1,LAMBDA(a,b,a+b))'),
    text('MAP({1,2},LAMBDA(x,ISREF(x)))'),
    text('MAP(D!B1:B2,LAMBDA(x,ISREF(x)))'),
    'MAP(D!B1:B3,LAMBDA(x,y,x))',
    'MAP(D!B1:B3,1)',
    text('MAP(D!E1:E4,LAMBDA(x,x*2))'),
    'MAP(D!B1:B2,LAMBDA(x,{1,2}))',
    text('MAP(D!B1:B3,ABS)'),
    text('MAP(D!B1:B2,D!C1:C2,SUM)'),
    text('MAP(D!A1:A3,LAMBDA(x,x))'),
    'MAP(5,LAMBDA(x,x+1))',
    text('MAP(D!B1:B3,LAMBDA(x,ROW(x)))'),
    'MAP(D!B1:B3,LAMBDA(x,D!C1:C3))',
    'MAP(1/0,LAMBDA(x,1))',
    text('MAP(D!B1:B2,LAMBDA(x,LAMBDA(y,y)))'),
  ],
  REDUCE: [
    'REDUCE(0,D!B1:B8,LAMBDA(a,b,a+b))',
    'REDUCE(,{2,3},LAMBDA(a,b,a+b^2))',
    'REDUCE(,{2,3},LAMBDA(a,b,IF(ISOMITTED(a),"om",a&b)))',
    'REDUCE(1,{2,3},LAMBDA(a,b,a*b))',
    text('REDUCE(0,{1,2},LAMBDA(a,b,a+{1,10}))'),
    'REDUCE(0,D!E1:E4,LAMBDA(a,b,a+b))',
    'REDUCE(0,D!B1:B3,LAMBDA(a,a))',
    'REDUCE(0,D!B1:B3,SUM)',
    'REDUCE("",D!F1:F3,LAMBDA(a,b,a&b))',
    'ISREF(REDUCE(D!B1,D!B1:B2,LAMBDA(a,b,b)))',
    'REDUCE(0,D!K1:K3,LAMBDA(a,b,a+ISBLANK(b)))',
  ],
  SCAN: [
    text('SCAN(0,D!B1:B4,LAMBDA(a,b,a+b))'),
    text('SCAN(,{1,2,3},LAMBDA(a,b,a+b))'),
    text('SCAN(0,{1,2;3,4},LAMBDA(a,b,a+b))'),
    'SCAN(0,{1,2},LAMBDA(a,b,{1,2}))',
    text('SCAN("",{"a","b"},LAMBDA(a,b,a&b))'),
    text('SCAN(0,D!E1:E4,LAMBDA(a,b,a+b))'),
  ],
  BYROW: [
    text('BYROW(D!B1:C3,LAMBDA(r,SUM(r)))'),
    text('BYROW(D!B1:C3,SUM)'),
    text('BYROW(D!B1:C3,LAMBDA(r,ISREF(r)))'),
    text('BYROW({1,2;3,4},LAMBDA(r,COLUMNS(r)))'),
    'BYROW(D!B1:C3,LAMBDA(r,r))',
    text('BYROW(D!B1:B3,LAMBDA(r,r))'),
    'BYROW(D!B1:C3,LAMBDA(a,b,a))',
    'BYROW(5,LAMBDA(r,r+1))',
    text('BYROW(D!B1:C2,LAMBDA(r,ROW(r)))'),
    'BYROW(NA(),LAMBDA(r,1))',
  ],
  BYCOL: [text('BYCOL(D!B1:C3,MAX)'), text('BYCOL(D!E1:E4,SUM)'), text('BYCOL(D!B1:C3,LAMBDA(c,ROWS(c)))'), 'BYCOL(D!B1:C3,LAMBDA(c,c))'],
  MAKEARRAY: [
    text('MAKEARRAY(2,3,LAMBDA(r,c,r*10+c))'),
    'MAKEARRAY(0,2,LAMBDA(r,c,1))',
    text('MAKEARRAY(2.7,1.2,LAMBDA(r,c,r))'),
    text('MAKEARRAY("2",1,LAMBDA(r,c,r))'),
    'MAKEARRAY(2,2,LAMBDA(r,r))',
    'MAKEARRAY(2,2,LAMBDA(r,c,{1,2}))',
    text('MAKEARRAY(2,2,SUM)'),
    'MAKEARRAY(-1,1,LAMBDA(r,c,r))',
    text('MAKEARRAY(TRUE,1,LAMBDA(r,c,r))'),
    'MAKEARRAY(1,1,1)',
  ],
}

export function buildProbes(): Probe[] {
  const out: Probe[] = []
  for (const [fn, formulas] of Object.entries(CASES)) formulas.forEach((f, i) => out.push({ id: `${fn}#${i + 1}`, group: fn, fn, formula: xl(xlpm(f)) }))
  return out
}

/** The engine's answers by case id, committed before Excel ran the set. */
export const PREDICTED: Readonly<Record<string, Val>> = JSON.parse(readFileSync(new URL('./lambda.predicted.json', import.meta.url), 'utf8'))
