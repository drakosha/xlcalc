/**
 * Held-out cases for LET: each answer was predicted from the rules the `let` set found and committed before Excel
 * computed the case (`refused`: Excel does not load the formula). The rules: a name is seen from the end of its
 * value to the end of its LET, case aside, an inner LET shadowing; a value is computed once and stays a reference
 * when it is one; a left-out value is a blank; a call that ends with a value answers it; a name declared twice,
 * no name at all or more than 126 names are not loaded. Answers: `test/excel/cases/let-held.json`.
 */
import type { Val } from '../formulajs.js'
import type { Probe } from '../probes.js'
import { collector } from './arrays/common.js'
import { xlpm } from './let/xlpm.js'

const n = (v: number): Val => ({ t: 'n', v })
const s = (v: string): Val => ({ t: 's', v })
const b = (v: boolean): Val => ({ t: 'b', v })
const e = (v: Extract<Val, { t: 'e' }>['v']): Val => ({ t: 'e', v })

/** 127 names, each the previous one plus 1, and no calculation. */
const names127 = `LET(${Array.from({ length: 127 }, (_, i) => `n_${i + 1},${i ? `n_${i}+1` : '1'}`).join(',')})`

/** Formula (the engine's spelling) and the predicted answer. */
const CASES: [string, Val | 'refused'][] = [
  ['LET(x,5,y,x*2)', n(10)],
  ['LET(x,5,y,y)', e('NAME')],
  ['LET(x)', 'refused'],
  ['LET(x,1,X,2,x)', 'refused'],
  [names127, 'refused'],
  ['LET(x,,ISBLANK(x))', b(true)],
  ['LET(x,,x&"")', s('')],
  ['LET(x,,TYPE(x))', n(1)],
  ['LET(r,D!B1:B8,INDEX(r,3))', n(15)],
  ['LET(r,D!B1:C8,SUM(INDEX(r,0,2)))', n(108)],
  ['LET(a,D!B1:B3,b,a*2,SUM(b))', n(54)],
  ['LET(x,D!B1:B8,COUNTIF(x,">10"))', n(5)],
  ['LET(t,"abc",LEN(t)&UPPER(t))', s('3ABC')],
  ['LET(x,2,y,3,LET(x,y,z,x*10,z+x))', n(33)],
  ['LET(v,D!A2,v*1)', n(2)],
  ['LET(v,D!A2,COUNT(v))', n(0)],
  ['LET(v,D!A2,COUNTA(v))', n(1)],
  ['LET(v,{1,2;3,4},ROWS(v)*10+COLUMNS(v))', n(22)],
  ['LET(v,{1,2;3,4},INDEX(v,2,1))', n(3)],
  ['LET(x,TODAY(),x-x)', n(0)],
  ['LET(x,NA(),ISNA(x))', b(true)],
  ['LET(x,1/0,y,x,1)', n(1)],
  ['LET(x,D!A7,IFERROR(x,"e"))', s('e')],
  ['LET(f,SUM(D!B1:B8),f/COUNT(D!B1:B8))', n(14.34375)],
  ['SUM(LET(x,{1,2,3},x*x))', n(14)],
  ['LET(x,{1,2,3},SUM(x*x))', n(14)],
  ['LET(x,1,y,x+1,z,y+1,x*100+y*10+z)', n(123)],
  ['IF(LET(x,TRUE,x),"y","n")', s('y')],
  ['LET(_,5,_*2)', n(10)],
  ['LET(x,"5",x+1)', n(6)],
  ['LET(x,D!K1,x+1)', n(1)],
  ['LET(x,D!K1,ISBLANK(x))', b(true)],
  ['ROWS(LET(x,D!B1:C8,y,INDEX(x,0,1),y))', n(8)],
  ['LET(x,5,y,x*2)+1', n(11)],
  ['LET(x,D!B1:B8,y,SUM(x))', n(114.75)],
]

export function buildProbes(): Probe[] {
  const { out, add } = collector()
  for (const [f] of CASES) add('held', 'LET', xlpm(f))
  return out
}

/** The predicted answers by case id. */
export const PREDICTED: Readonly<Record<string, Val | 'refused'>> = Object.fromEntries(CASES.map(([, v], i) => [`held#${i + 1}`, v]))
