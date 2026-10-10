/**
 * The format Excel gives a typed formula by what it computes, measured in the entry sets `typed` and `typed-held`:
 * `TODAY()` and `DATE(…)` show the short date, `NOW()` the date and time, `TIME(…)` the time with AM or PM, `PMT(…)` and
 * its kin money. The format passes through a sign, parentheses, `SUM`, `MAX` and `MIN` (their first argument with one)
 * and a sum (its first term with one: `1+TODAY()`), and through a difference whose second term is a plain number
 * (`TODAY()-1`), not through other functions (`YEAR(TODAY())`, `EDATE(…)`, `IF(…)`) or other operators. A reference takes
 * its cell's format in Excel, which a typed entry without its book does not know: it gives none.
 */
import type { NumberFormat } from '../contract.js'
import type { Ast } from '../syntax.js'

/** The functions whose result has a format, by name; `money` is the locale's money with two decimals. */
const FORMATS: Readonly<Record<string, NumberFormat | 'money'>> = { TODAY: 14, DATE: 14, NOW: 22, TIME: 18, PMT: 'money', PV: 'money', FV: 'money', IPMT: 'money', PPMT: 'money', NPV: 'money' }

/** The functions that take the format of their first argument with one. */
const PASSING = new Set(['SUM', 'MAX', 'MIN'])

/** Whether a node is a number with no format of its own: literals and arithmetic of them. */
const plain = (a: Ast): boolean => a.k === 'num' || ((a.k === 'un' || a.k === 'paren' || a.k === 'pct') && plain(a.a)) || (a.k === 'bin' && plain(a.a) && plain(a.b))

/** The format a formula's result takes, `money` standing for the locale's; undefined when it has none. */
export function formulaFormat(a: Ast): NumberFormat | 'money' | undefined {
  switch (a.k) {
    case 'fn':
      if (FORMATS[a.name] !== undefined) return FORMATS[a.name]
      return PASSING.has(a.name) ? a.args.map(formulaFormat).find((f) => f !== undefined) : undefined
    case 'un':
    case 'paren':
      return formulaFormat(a.a)
    case 'bin':
      if (a.op === '+') return formulaFormat(a.a) ?? formulaFormat(a.b)
      return a.op === '-' && plain(a.b) ? formulaFormat(a.a) : undefined
    default:
      return undefined
  }
}
