/**
 * An omitted argument in every whitelisted function: its typical call (`arrays/lift.ts`, the functions with an
 * array answer from `NEWER` inside `ARRAYTOTEXT(…,1)`, which shows the whole array, the fitted numbers rounded to
 * 6 decimals: computed by another algorithm, their text can differ in the 15th digit) with each argument left out
 * in turn and, beside it, replaced by 0, "" and FALSE; and, where the function takes one more argument, the call
 * itself, with that argument left out and given as 0, "" and FALSE. The answers tell, function by function,
 * whether a left-out argument is a blank (read as 0, "" or FALSE by its kind) or the argument not given. Positions
 * Excel takes only as a reference are left to the held-out set; size arguments get no text or boolean.
 */
import { FUNCTIONS } from '../../../../src/functions.js'
import type { Add } from '../arrays/common.js'
import { splitArgs, typicalCalls } from '../arrays/lift.js'
import { isRefOnly, isSize } from '../fuzz/pools.js'

/** Typical calls of the functions with an array answer that the measurement has none for (RANDARRAY: `omit/kinds.ts`). */
const NEWER: Record<string, string> = {
  CHOOSECOLS: 'CHOOSECOLS(D!B1:C3,2,1)',
  CHOOSEROWS: 'CHOOSEROWS(D!B1:C3,3,1)',
  DROP: 'DROP(D!B1:C3,1,1)',
  EXPAND: 'EXPAND({1,2},2,3,0)',
  FILTER: 'FILTER(D!B1:B8,D!B1:B8>20,"none")',
  FREQUENCY: 'FREQUENCY(D!B1:B8,D!AT1:AT3)',
  GROWTH: 'GROWTH(D!BB1:BB6,D!BC1:BC6,D!BD1:BD2,TRUE)',
  HSTACK: 'HSTACK(1,D!B1:B2,"a")',
  LINEST: 'LINEST(D!AY1:AY5,D!AZ1:AZ5,TRUE,FALSE)',
  LOGEST: 'LOGEST(D!BB1:BB6,D!BC1:BC6,TRUE,FALSE)',
  MMULT: 'MMULT(D!P1:Q2,D!R1:S2)',
  'MODE.MULT': 'MODE.MULT(1,2,2,3,3)',
  SEQUENCE: 'SEQUENCE(2,3,10,5)',
  SORT: 'SORT(D!B1:C4,2,-1,FALSE)',
  SORTBY: 'SORTBY(D!B1:B4,D!C1:C4,-1)',
  TAKE: 'TAKE(D!B1:C4,2,1)',
  TEXTSPLIT: 'TEXTSPLIT("a,b;c,,d",",",";",TRUE,0,"-")',
  TOCOL: 'TOCOL(D!B1:C2,0,TRUE)',
  TOROW: 'TOROW(D!B1:C2,0,TRUE)',
  TREND: 'TREND(D!AY1:AY5,D!AZ1:AZ5,D!BA1:BA2,TRUE)',
  UNIQUE: 'UNIQUE(D!AS1:AS8,FALSE,TRUE)',
  VSTACK: 'VSTACK(1,D!B1:B2,"a")',
  WRAPCOLS: 'WRAPCOLS({1,2,3},2,0)',
  WRAPROWS: 'WRAPROWS({1,2,3},2,0)',
}

/** Functions whose numbers come from a fit: rounded before they become text. */
const FITTED = new Set(['GROWTH', 'LINEST', 'LOGEST', 'TREND'])

const newer = (): [string, string[]][] => Object.entries(NEWER).map(([fn, text]) => [fn, splitArgs(text.slice(fn.length + 1, -1))])

export function all(add: Add) {
  for (const [fn, args] of [...typicalCalls(), ...newer()]) {
    const text = (xs: string[]) => `${fn}(${xs.join(',')})`
    const call = (xs: string[]) => (FITTED.has(fn) ? `ARRAYTOTEXT(ROUND(${text(xs)},6),1)` : fn in NEWER ? `ARRAYTOTEXT(${text(xs)},1)` : text(xs))
    const g = (...fs: string[]) => add(`all:${fn}`, fn, ...fs)
    const variants = (i: number) => (isSize(fn, i) ? ['', '0'] : ['', '0', '""', 'FALSE'])
    const k = args.length
    // a call of one argument left out is a call of none, a different count
    args.forEach((_, i) => {
      if (k > 1 && !isRefOnly(fn, i)) g(...variants(i).map((v) => call(args.map((x, j) => (j === i ? v : x)))))
    })
    if (FUNCTIONS[fn].takes(k + 1) && !isRefOnly(fn, k)) g(call(args), ...variants(k).map((v) => call([...args, v])))
  }
}
