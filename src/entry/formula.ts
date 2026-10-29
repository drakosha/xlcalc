/**
 * The format Excel gives a typed formula by what it computes, measured in the entry sets `typed`, `typed-held`,
 * `typed-held-2`, `refs`, `refs-held` and `refs-held-2`: `TODAY()` and `DATE(…)` show the short date, `NOW()` the date
 * and time, `TIME(…)` the time with AM or PM, `PMT(…)` and its kin money, a reference to the book's own cells the format
 * of its first cell (`at`: C1 for `C:C`, none for General). The format passes through a sign, parentheses, `SUM`, `MAX`,
 * `MIN`, `AVERAGE`, `MEDIAN`, `ROUND`, `ROUNDUP`, `ROUNDDOWN`, `TRUNC`, `INT` and `MOD` (their first argument with
 * one), a sum (its first term with one: `1+TODAY()`) and a difference (its first term with one, else its second; none
 * for two dates, a number of days), not through other functions (`ABS(…)`, `MROUND(…)`, `IF(…)`, `YEAR(…)`) or other
 * operators (`*`, `/`, `&`). A name and a link to another book give none.
 */
import type { NumberFormat } from '../contract.js'
import type { Ast } from '../syntax.js'
import { shows } from './forms.js'

/** The format a cell shows, its own or its row's or column's, by its sheet (null: the formula's own) and its zero-based row and column. */
export type FormatAt = (sheet: string | null, row: number, col: number) => NumberFormat | undefined

const general = (f: NumberFormat | undefined) => f === undefined || f === 0 || (typeof f === 'string' && f.toLowerCase() === 'general')

/** The functions whose result has a format, by name; `money` is the locale's money with two decimals. */
const FORMATS: Readonly<Record<string, NumberFormat | 'money'>> = { TODAY: 14, DATE: 14, NOW: 22, TIME: 18, PMT: 'money', PV: 'money', FV: 'money', IPMT: 'money', PPMT: 'money', NPV: 'money' }

/** The functions that take the format of their first argument with one. */
const PASSING = new Set(['SUM', 'MAX', 'MIN', 'AVERAGE', 'MEDIAN', 'ROUND', 'ROUNDUP', 'ROUNDDOWN', 'TRUNC', 'INT', 'MOD'])

/** Whether a formula's format shows a date. */
const isDate = (f: NumberFormat | 'money' | undefined) => f !== undefined && f !== 'money' && shows(f, 'day')

/** The format a formula's result takes, `money` standing for the locale's; undefined when it has none. */
export function formulaFormat(a: Ast, at?: FormatAt): NumberFormat | 'money' | undefined {
  const of = (x: Ast) => formulaFormat(x, at)
  switch (a.k) {
    case 'ref': {
      const f = at?.(a.sheet, a.ref.r1, a.ref.c1)
      return general(f) ? undefined : f
    }
    case 'fn':
      if (FORMATS[a.name] !== undefined) return FORMATS[a.name]
      return PASSING.has(a.name) ? a.args.map(of).find((f) => f !== undefined) : undefined
    case 'un':
    case 'paren':
      return of(a.a)
    case 'bin':
      if (a.op === '+') return of(a.a) ?? of(a.b)
      if (a.op !== '-') return undefined
      // a difference of two dates is a number of days
      const [x, y] = [of(a.a), of(a.b)]
      return isDate(x) && isDate(y) ? undefined : (x ?? y)
    default:
      return undefined
  }
}
