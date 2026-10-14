/**
 * Held out: a negative number to a fractional power, which the engine answered `#NUM!` while Excel answers the real
 * root where the power is the reciprocal of an odd whole number (`(-32)^(1/5)` is -2, `(-8)^(2/3)` `#NUM!`, from a
 * user's model). The rule predicted: where 1/y, to 15 significant digits, is an odd whole number, x^y is -(-x)^y;
 * otherwise `#NUM!` as before. The cases: 1/n for odd and even n, its negative, p/q with an odd q, powers a little off
 * 1/3, reciprocals whose 1/y is not n in doubles (1/49, 1/1000001, 1/999999999999999) or an odd n past 15 digits
 * (2^52+1), bases -0, -1, ±1E±300, the operator and POWER alike, cells, arrays and SERIESSUM. P!A1 is -8, P!A2 `1/3`.
 * Answers: `test/excel/cases/power-neg-held.json`; `PREDICTED` was committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import type { Probe, SheetCells } from '../probes.js'
import { heldGroups } from './arrays/common.js'
import { n } from '../vals.js'

const NUM: Val = { t: 'e', v: 'NUM' }
/** -(-x)^y, the real root predicted. */
const real = (x: number, y: number) => -Math.pow(-x, y)
const root = (x: number, y: number) => n(real(x, y))

export const SHEETS: SheetCells = { P: { A1: -8, A2: { formula: '1/3' } } }

/** Group, function, formula and the answer predicted before Excel ran. */
const CASES: [string, string, string, Val][] = [
  // the user's: 1/5 and 0.2, 1/3, a small base; 2/3 and 1/2
  ['caret', '^', '(-32)^(1/5)', n(-2)],
  ['caret', '^', '(-32)^0.2', n(-2)],
  ['caret', '^', '(-8)^(1/3)', n(-2)],
  ['caret', '^', '(-1E-10)^(1/5)', root(-1e-10, 1 / 5)],
  ['caret', '^', '(-8)^(2/3)', NUM],
  ['caret', '^', '(-8)^(1/2)', NUM],
  // 1/n for odd and even n
  ['caret', '^', '(-2)^(1/3)', root(-2, 1 / 3)],
  ['caret', '^', '(-2)^(1/5)', root(-2, 1 / 5)],
  ['caret', '^', '(-2)^(1/7)', root(-2, 1 / 7)],
  ['caret', '^', '(-2)^(1/9)', root(-2, 1 / 9)],
  ['caret', '^', '(-2)^(1/2)', NUM],
  ['caret', '^', '(-2)^(1/4)', NUM],
  ['caret', '^', '(-2)^(1/101)', root(-2, 1 / 101)],
  ['caret', '^', '(-2)^(1/1000001)', root(-2, 1 / 1000001)],
  ['caret', '^', '(-2)^(1/1000000)', NUM],
  // a negative reciprocal; p/q with an odd q
  ['caret', '^', '(-8)^(-1/3)', root(-8, -1 / 3)],
  ['caret', '^', '(-8)^(-1/5)', root(-8, -1 / 5)],
  ['caret', '^', '(-32)^(3/5)', NUM],
  ['caret', '^', '(-32)^0.6', NUM],
  ['caret', '^', '(-8)^(-2/3)', NUM],
  // a little off 1/3: the same double, 14 digits, one or two units in the last place either way
  ['caret', '^', '(-8)^0.3333333333333333', n(-2)],
  ['caret', '^', '(-8)^0.33333333333333', NUM],
  ['caret', '^', '(-8)^(1/3+1E-16)', root(-8, 1 / 3 + 1e-16)],
  ['caret', '^', '(-8)^(1/3-1E-16)', root(-8, 1 / 3 - 1e-16)],
  // 1/y not n in doubles (1/(1/49) is 49.00000000000001), an odd n past 15 digits, 1/y beyond whole numbers
  ['caret', '^', '(-8)^(1/49)', root(-8, 1 / 49)],
  ['caret', '^', '(-2)^(1/999999999999999)', root(-2, 1 / 999999999999999)],
  ['caret', '^', '(-2)^(1/(2^52+1))', NUM],
  ['caret', '^', '(-2)^1E-20', NUM],
  // bases -0, -1, ±1E±300
  ['caret', '^', '(-0)^(1/3)', n(0)],
  ['caret', '^', '(-1)^(1/3)', n(-1)],
  ['caret', '^', '(-1E-300)^(1/3)', root(-1e-300, 1 / 3)],
  ['caret', '^', '(-1E+300)^(1/3)', root(-1e300, 1 / 3)],
  ['caret', '^', '(-1E-300)^(1/101)', root(-1e-300, 1 / 101)],
  ['caret', '^', '(-1E+300)^(1/2)', NUM],
  // the minus binds first; a whole power stays
  ['caret', '^', '-8^(1/3)', n(-2)],
  ['caret', '^', '(-8)^(1/1)', n(-8)],
  // cells and arrays
  ['caret', '^', 'P!A1^(1/3)', n(-2)],
  ['caret', '^', 'P!A1^P!A2', n(-2)],
  ['caret', '^', 'INDEX({-8,-27,-32}^(1/3),2)', root(-27, 1 / 3)],
  ['caret', '^', 'INDEX((-8)^(1/{3,5,2}),3)', NUM],
  ['caret', '^', 'INDEX((-8)^(1/{3,5,2}),1)', n(-2)],
  ['caret', '^', 'SUM({-8;-27}^(1/3))', n(-2 + real(-27, 1 / 3))],
  // POWER as the operator
  ['POWER', 'POWER', 'POWER(-32,0.2)', n(-2)],
  ['POWER', 'POWER', 'POWER(-8,1/3)', n(-2)],
  ['POWER', 'POWER', 'POWER(-8,2/3)', NUM],
  ['POWER', 'POWER', 'POWER(-8,1/2)', NUM],
  ['POWER', 'POWER', 'POWER(-8,-1/3)', root(-8, -1 / 3)],
  ['POWER', 'POWER', 'POWER(-8,1/49)', root(-8, 1 / 49)],
  ['POWER', 'POWER', 'POWER(-1E-10,1/5)', root(-1e-10, 1 / 5)],
  ['POWER', 'POWER', 'POWER(-8,0.33333333333333)', NUM],
  ['POWER', 'POWER', 'POWER(-2,1/999999999999999)', root(-2, 1 / 999999999999999)],
  ['POWER', 'POWER', 'POWER(-0,1/3)', n(0)],
  ['POWER', 'POWER', 'SUM(POWER({-8;-27},1/3))', n(-2 + real(-27, 1 / 3))],
  ['POWER', 'POWER', 'POWER(P!A1,P!A2)', n(-2)],
  // SERIESSUM adds 0 for a negative x to a fractional power (measured before), predicted to stay so
  ['SERIESSUM', 'SERIESSUM', 'SERIESSUM(-8,1/3,1,{1})', n(0)],
  ['SERIESSUM', 'SERIESSUM', 'SERIESSUM(-8,1/2,1,{1})', n(0)],
]

/** The set's cases and the answers predicted, by case id. */
export const { buildProbes, PREDICTED } = heldGroups(CASES)

/**
 * The predictions Excel did not confirm (the Mac and Windows alike): 1/y of 999999999999999 is too large for a root,
 * `#NUM!` (`power-edges-held` and `power-bound-held` found the bound below 2^31-1). The roots themselves are Excel's
 * -exp(y·ln(-x)), whose last digit the prediction's -(-x)^y missed within the tolerance (`(-8)^(1/3)` -1.9999999999999998).
 */
export const MISSED: readonly string[] = ['caret#26', 'POWER#9']
