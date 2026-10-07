/**
 * Excel cases for LET: names over values, arrays and references, scope (later names see earlier ones, an inner
 * LET shadows an outer one, a name is no longer seen past its LET), names that look like functions, the case of
 * names, LET inside functions and functions inside LET, how often a value is computed, errors in values, an
 * array result, and the calls Excel does or does not load. The names are written in Excel's file spelling
 * (`_xlpm.x`, `let/xlpm.ts`). Answers: `test/excel/cases/let.json`.
 */
import type { Probe } from '../probes.js'
import { collector, xl } from './arrays/common.js'
import { xlpm } from './let/xlpm.js'

/** `count` names `n_1`…, each the previous one plus 1; the last one is the result. */
const pairs = (count: number) => `LET(${Array.from({ length: count }, (_, i) => `n_${i + 1},${i ? `n_${i}+1` : '1'}`).join(',')},n_${count})`

/** Calls Excel does not load (a book with one is refused): a name declared twice, a reference, a number or text as a name, 127 names. */
export const REFUSED: readonly string[] = ['LET(_xlpm.x,1,_xlpm.x,2,_xlpm.x)', 'LET(A1,1,A1)', 'LET(1,2,3)', 'LET("x",1,2)', xlpm(pairs(127))].map(xl)

export function buildProbes(): Probe[] {
  const { out, add } = collector()
  const g = (group: string, ...fs: string[]) => add(group, 'LET', ...fs.map(xlpm))
  g('let:value', 'LET(x,1,x+1)', 'LET(x,2,y,3,x*y)', 'LET(x,1,y,x+1,z,y*10,z+x)', 'LET(x,"a",x&"b")', 'LET(x,TRUE,NOT(x))', 'LET(x,D!B2,x*2)', 'LET(x,5,x%)', 'LET(x,-2,-x^2)', 'LET(x,1,y,2,x)')
  g('let:array', 'LET(a,{1,2,3},SUM(a))', 'LET(a,{1,2,3},a*2)', 'LET(a,{1,2,3},INDEX(a,2))', 'LET(a,{1},TYPE(a))', 'TYPE(LET(a,{1},a))', 'LET(x,SEQUENCE(3),ROWS(x))', 'LET(x,SEQUENCE(3),x)', 'LET(c,{TRUE,FALSE},SUM(IF(c,1,2)))', 'LET(x,D!B1:B8,SUM(x*2))', 'SUM(LET(x,D!B1:B8,x*2))')
  // a reference stays a reference: SUM skips the text "2" of D!A2 and the text of D!D1:D8
  g('let:ref', 'LET(r,D!B1:B8,ROWS(r))', 'LET(r,D!D1:D8,SUM(r))', 'LET(r,D!A2,SUM(r))', 'LET(r,D!A2,r+1)', 'LET(r,D!B1:B8,SUMIF(r,">10"))', 'LET(r,D!B1:B8,r)', 'LET(r,D!B1:B8,COLUMN(r))')
  g('let:ref', 'LET(r,D!A1,ISBLANK(r))', 'LET(r,D!A1,r)', 'LET(r,D!A1,r&"")', 'LET(r,D!B1:B2,TYPE(r))', 'LET(r,D!B1,TYPE(r))', 'LET(r,INDEX(D!B1:C8,0,1),SUM(r))', 'LET(r,IF(TRUE,D!A2),SUM(r))', 'LET(r,D!K1:K8,COUNTA(r))')
  g('let:result', 'SUM(LET(r,D!D1:D8,r))', 'SUMIF(LET(r,D!B1:B8,r),">10")', 'ROWS(LET(x,D!B1:B8,x))', 'COUNTA(LET(r,D!K1:K8,r))', 'SUM(LET(r,D!A2,r))', 'LET(r,D!A2,r)')
  g('let:scope', 'LET(X,1,x+1)', 'LET(abc,2,ABC*Abc)', 'LET(x,1,LET(x,2,x)+x)', 'LET(x,1,y,LET(x,10,x+1),x+y)', 'LET(x,2,LET(y,x*3,y+x))', 'LET(x,1,x)+LET(x,2,x)')
  g('let:nest', 'SUM(LET(x,1,x),2)', 'IF(TRUE,LET(x,5,x),0)', 'LET(x,ROUND(2.567,1),x*2)', 'LET(x,D!B1:B8,MAX(x)-MIN(x))', 'LET(x,3,CHOOSE(x,"a","b","c"))', 'IFERROR(LET(x,1/0,x),"e")')
  g('let:names', 'LET(sum,5,SUM(sum,1))', 'LET(pi,3,pi*2+PI())', 'LET(_x,1,_x+1)', 'LET(x.y,1,x.y+1)', 'LET(x_1,2,x_1*2)', 'LET(икс,1,икс+1)', 'LET(if,1,IF(if,2,3))')
  // a value is computed once: RAND gives the same number to every use; an unused value's error does not count
  g('let:once', 'LET(x,RAND(),x-x)', 'LET(x,RAND(),x=x)', 'LET(x,1/0,1)', 'LET(x,NA(),y,2,y)', 'LET(x,1/0,x+1)', 'LET(x,1/0,IFERROR(x,7))', 'LET(x,1,y,1/0,x)')
  g('let:count', pairs(126))
  // a name used outside its LET (a defined name, none here), in its own value or before it is declared
  add('let:unbound', 'LET', 'LET(_xlpm.x,1,_xlpm.x)+x', 'LET(_xlpm.x,x+1,_xlpm.x)', 'LET(_xlpm.x,_xlpm.x+1,_xlpm.x)', 'LET(_xlpm.x,_xlpm.y,_xlpm.y,1,_xlpm.x)')
  // an even number of arguments: Excel loads it
  g('let:even', 'LET(x,1)', 'LET(x,1,y,2)')
  // in the file any word is a name after _xlpm., one that reads as a reference or a boolean too
  add('let:file', 'LET', 'LET(_xlpm.A1,1,_xlpm.A1)', 'LET(_xlpm.r,1,_xlpm.r)', 'LET(_xlpm.c,1,_xlpm.c)', 'LET(_xlpm.R1C1,1,_xlpm.R1C1)', 'LET(_xlpm.TRUE,1,_xlpm.TRUE)')
  g('let:omitted', 'LET(x,,x)', 'LET(x,,x+1)', 'LET(x,1,)', 'LET(x,1,y,,x)')
  return out
}
