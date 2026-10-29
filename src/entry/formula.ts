/**
 * The format Excel gives a typed formula by what it computes, measured in the entry sets `typed`, `typed-held` and
 * `typed-held-2`: `TODAY()` and `DATE(…)` show the short date, `NOW()` the date and time, `TIME(…)` the time with AM or
 * PM, `PMT(…)` and its kin money. The format passes through a sign, parentheses, `SUM`, `MAX`, `MIN`, `AVERAGE` and
 * `ROUND` (their first argument with one), a sum (its first term with one: `1+TODAY()`) and a difference whose second
 * term is a plain number (`TODAY()-1`), not through other functions (`YEAR(TODAY())`, `EDATE(…)`, `IF(…)`) or other
 * operators. A reference to the book's own cells takes the format of its first cell (`at`), a General one none; without
 * `at`, a link to another book and a name give none.
 */
import type { NumberFormat } from '../contract.js'
import type { Ast } from '../syntax.js'

/** The format a cell shows, its own or its row's or column's, by its sheet (null: the formula's own) and its zero-based row and column. */
export type FormatAt = (sheet: string | null, row: number, col: number) => NumberFormat | undefined

const general = (f: NumberFormat | undefined) => f === undefined || f === 0 || (typeof f === 'string' && f.toLowerCase() === 'general')

/** The functions whose result has a format, by name; `money` is the locale's money with two decimals. */
const FORMATS: Readonly<Record<string, NumberFormat | 'money'>> = { TODAY: 14, DATE: 14, NOW: 22, TIME: 18, PMT: 'money', PV: 'money', FV: 'money', IPMT: 'money', PPMT: 'money', NPV: 'money' }

/** The functions that take the format of their first argument with one. */
const PASSING = new Set(['SUM', 'MAX', 'MIN', 'AVERAGE', 'ROUND'])

/** Whether a node is a number with no format of its own: literals and arithmetic of them. */
const plain = (a: Ast): boolean => a.k === 'num' || ((a.k === 'un' || a.k === 'paren' || a.k === 'pct') && plain(a.a)) || (a.k === 'bin' && plain(a.a) && plain(a.b))

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
      return a.op === '-' && plain(a.b) ? of(a.a) : undefined
    default:
      return undefined
  }
}
