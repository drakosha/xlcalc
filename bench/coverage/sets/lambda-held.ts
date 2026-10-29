/**
 * Held out after the set `lambda`, its rules predicted: a LAMBDA takes exactly its parameters (those left out between
 * commas omitted); a function is a value of its own (TYPE 128, no error, an operator's and a function's #VALUE!,
 * ARRAYTOTEXT and VALUETOTEXT λ, a cell #CALC!); a function by its name called with counts it does not take is #CALC!;
 * an error given as an array is the answer; REDUCE and SCAN start from the first value without an initial one; MAKEARRAY
 * takes whole sizes. Answers: `test/excel/cases/lambda-held.json`; `PREDICTED`, the engine's answers
 * (`sets/predict.ts lambda-held`), was committed before Excel ran the set.
 */
import { readFileSync } from 'node:fs'
import type { Val } from '../formulajs.js'
import type { Probe } from '../probes.js'
import { xl } from './arrays/common.js'
import { xlpm } from './let/xlpm.js'

const t = (f: string, strict = 1) => `ARRAYTOTEXT(${f},${strict})`

const CASES: Record<string, string[]> = {
  LAMBDA: [
    'LAMBDA(x,y,z,1)(1,2)',
    'LAMBDA(x,y,1)(,)',
    'LAMBDA(x,y,TYPE(x)+TYPE(y))(,)',
    'TYPE(SUM)',
    'ISTEXT(LAMBDA(x,x))',
    'ISNONTEXT(LAMBDA(x,x))',
    'ISLOGICAL(SUM)',
    'LAMBDA(x,x)=1',
    'LAMBDA(x,x)&"a"',
    'IFNA(LAMBDA(x,x),1)',
    'CHOOSE(1,LAMBDA(x,x))',
    'COUNT(LAMBDA(x,x))',
    'COUNTA(LAMBDA(x,x),1)',
    t('LAMBDA(x,x)', 0),
    'VALUETOTEXT(LAMBDA(x,x))',
    'ROW(LAMBDA(r,r)(D!B5))',
    'INDEX(MAP({1,2},LAMBDA(x,LAMBDA(y,y))),1)',
    t('BYROW({1;2},LAMBDA(r,LAMBDA(x,x)))'),
    'LAMBDA(f,f(3))(LAMBDA(x,x*5))',
    'LAMBDA(f,f(3))(ABS)',
    'LAMBDA(n,IF(n>2,"big","small"))(D!B1)',
    'SUM(LAMBDA(r,r*2)(D!B1:B3))',
  ],
  SUM: ['LET(f,SUM,f(1,2,3))', 'IF(TRUE,SUM)(4,5)'],
  ABS: ['LET(f,ABS,f(1,2))'],
  ROUND: [t('MAP(D!B1:B2,ROUND)'), t('MAP({1.26,2.71},{1,0},ROUND)')],
  MAP: [
    'MAP(D!B1:B2,1/0,LAMBDA(a,b,a))',
    t('MAP(D!B1:B3,LAMBDA(x,REDUCE(0,SEQUENCE(MIN(x,5)),LAMBDA(a,b,a+b))))'),
    t('MAP(D!K1:K2,LAMBDA(x,ISBLANK(x)))'),
    t('MAP({1,2;3,4},LAMBDA(x,x*10))'),
    t('MAP(D!B1:B2,D!C1:C3,LAMBDA(a,b,ISNA(b)))'),
  ],
  REDUCE: [
    'REDUCE(,{"a","b","c"},LAMBDA(a,b,a&b))',
    'REDUCE(,5,LAMBDA(a,b,a+b))',
    'REDUCE(,{1,2},LAMBDA(a,b,ISOMITTED(a)))',
    'REDUCE(,D!B1:B3,LAMBDA(a,b,ISREF(a)))',
    'REDUCE(1/0,{1,2},LAMBDA(a,b,ISERROR(a)))',
  ],
  SCAN: [t('SCAN(,{1,2},LAMBDA(a,b,a*10+b))'), t('SCAN(,{5},LAMBDA(a,b,a+b))'), t('SCAN(,D!B1:B3,LAMBDA(a,b,a+b))')],
  MAKEARRAY: ['MAKEARRAY(2.0000001,1,LAMBDA(r,c,r))', 'MAKEARRAY("2.5",1,LAMBDA(r,c,r))', t('MAKEARRAY(1,3,LAMBDA(r,c,c))'), 'MAKEARRAY(1/0,1,LAMBDA(r,c,r))'],
  BYCOL: [t('BYCOL({1,2;3,4},LAMBDA(c,SUM(c)))'), 'BYCOL(D!B1:C2,LAMBDA(a,b,1))'],
  ISOMITTED: ['LAMBDA(x,y,ISOMITTED(x)+ISOMITTED(y))(,)', 'MAP({1},LAMBDA(x,ISOMITTED(x)))'],
}

export function buildProbes(): Probe[] {
  const out: Probe[] = []
  for (const [fn, formulas] of Object.entries(CASES)) formulas.forEach((f, i) => out.push({ id: `${fn}#${i + 1}`, group: fn, fn, formula: xl(xlpm(f)) }))
  return out
}

/** The engine's answers by case id, committed before Excel ran the set. */
export const PREDICTED: Readonly<Record<string, Val>> = JSON.parse(readFileSync(new URL('./lambda-held.predicted.json', import.meta.url), 'utf8'))

/**
 * The predictions Excel for Mac did not confirm: a function compares with nothing (`LAMBDA(x,x)=1` is #VALUE!), COUNT
 * passes it by, a function by its name called with more arguments than it takes is #CALC! too, MAP's later array an
 * error is #VALUE!, and a position a smaller array lacks is #N/A without a call.
 */
export const MISSED: readonly string[] = ['LAMBDA#8', 'LAMBDA#12', 'ABS#1', 'MAP#1', 'MAP#5']
