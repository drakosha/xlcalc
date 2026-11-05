/**
 * The fourth round of `file-refusals`: every function of the whitelist one argument short of its least and one over its most, for
 * the writer's check of counts holds only where Excel's table is the library's. And how many tokens a formula holds, since the third
 * round's 8,192 characters of references were refused where 8,192 of calls were not.
 *
 * `refused` and `predicted` were committed before Excel ran this set, with the Windows answers of the other rounds in hand. The
 * model: Excel's counts are the library's, but where the first round showed otherwise (`BYROW` and `BYCOL` of one) and for the
 * functions Excel for Windows 16.0.20430 does not have, which it reads as a user-defined call. A formula holds up to 16,384 bytes
 * of tokens at BIFF12's sizes: 7 a cell, 13 a range, 1 an operator.
 */
import { arityOf } from '../../../src/arity.js'
import { FUNCTION_ARITY } from '../../../src/function-arity.js'
import { XLFN, XLWS } from '../../../src/xlsx/formulas/prefixes.js'
import type { LinksCase, LinksCases } from './cases.js'
import { cell, list, REFUSED } from './file-refusals.js'

/** What this Excel does not have (the set file-refusals-2) and reads as a call of a user's function; `BYROW`, `BYCOL` as measured. */
const UNREFUSED = new Set(['DETECTLANGUAGE', 'EUROCONVERT', 'GROUPBY', 'JIS', 'PERCENTOF', 'PIVOTBY', 'REGEXEXTRACT', 'REGEXREPLACE', 'REGEXTEST', 'TRANSLATE'])
const FEW_TAKEN = new Set(['BYROW', 'BYCOL'])

const spelled = (fn: string) => (XLWS.has(fn) ? `_xlfn._xlws.${fn}` : XLFN.has(fn) ? `_xlfn.${fn}` : fn)
const call = (fn: string, n: number) => `${spelled(fn)}(${list(n, () => 'B1', ',')})`
const idOf = (fn: string) => fn.toLowerCase().replaceAll('.', '-')

const counts = (): LinksCase[] =>
  Object.keys(FUNCTION_ARITY)
    .sort()
    .flatMap((fn) => {
      const { min, max } = arityOf(FUNCTION_ARITY[fn])
      const taken = UNREFUSED.has(fn)
      return [
        ...(min > 0 ? [cell(`few-${idOf(fn)}`, `${fn} of ${min - 1}`, call(fn, min - 1), taken || FEW_TAKEN.has(fn) ? {} : REFUSED)] : []),
        ...(max < 255 ? [cell(`many-${idOf(fn)}`, `${fn} of ${max + 1}`, call(fn, max + 1), taken ? {} : REFUSED)] : []),
      ]
    })

const terms = (unit: string, n: number) => list(n, () => unit, '+')

const tokens = (): LinksCase[] => [
  cell('cells-2048', '2,048 cells joined, 16,383 bytes of tokens', terms('B1', 2048), { value: 2048 }),
  cell('cells-2049', '2,049 cells joined, 16,391 bytes of tokens', terms('B1', 2049), REFUSED),
  cell('cells-2100', '2,100 cells joined', terms('B1', 2100), REFUSED),
  cell('ranges-1170', '1,170 ranges in SUM by pairs, about 16,380 bytes of tokens', terms('SUM(B1:B2)', 800)),
  cell('sum-cells', 'eight SUM of 255 cells, 14,360 bytes of tokens', terms(`SUM(${list(255, () => 'B1', ',')})`, 8), { value: 2040 }),
  cell('sum-cells-over', 'ten SUM of 255 cells, 17,950 bytes of tokens', terms(`SUM(${list(255, () => 'B1', ',')})`, 10), REFUSED),
]

export function buildLinksCases(): LinksCases {
  return { cases: [...counts(), ...tokens()] }
}
