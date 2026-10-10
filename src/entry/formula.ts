/**
 * The format Excel gives a formula typed into a cell in General by what it computes (the entry set `typed`: `TODAY()`
 * and `DATE(…)` show the short date, `NOW()` the date and time, `TIME(…)` the time with AM or PM, `PMT(…)` money): the
 * first call of such a function in the formula, read left to right, sets it. A reference takes its cell's format in
 * Excel, which a typed entry without its book does not know.
 */
import type { NumberFormat } from '../contract.js'
import type { Ast } from '../syntax.js'

/** The functions whose result sets a format, by name; `money` is the locale's money with two decimals. */
const FORMATS: Readonly<Record<string, NumberFormat | 'money'>> = {
  TODAY: 14,
  DATE: 14,
  DATEVALUE: 14,
  EDATE: 14,
  EOMONTH: 14,
  WORKDAY: 14,
  'WORKDAY.INTL': 14,
  NOW: 22,
  TIME: 18,
  TIMEVALUE: 18,
  PMT: 'money',
  PV: 'money',
  FV: 'money',
  IPMT: 'money',
  PPMT: 'money',
  NPV: 'money',
}

/** The children of a node, left to right. */
function children(a: Ast): Ast[] {
  switch (a.k) {
    case 'fn':
      return a.args
    case 'un':
    case 'pct':
    case 'paren':
    case 'at':
      return [a.a]
    case 'bin':
      return [a.a, a.b]
    case 'union':
      return a.items
    case 'let':
      return [...a.binds.map((b) => b.value), ...(a.body ? [a.body] : [])]
    default:
      return []
  }
}

/** The format the formula's first formatting call sets, `money` standing for the locale's; undefined when none does. */
export function formulaFormat(a: Ast): NumberFormat | 'money' | undefined {
  if (a.k === 'fn' && FORMATS[a.name] !== undefined) return FORMATS[a.name]
  for (const c of children(a)) {
    const f = formulaFormat(c)
    if (f !== undefined) return f
  }
  return undefined
}
