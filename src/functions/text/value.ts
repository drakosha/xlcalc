/** VALUE and NUMBERVALUE, this engine's own: text read by the book's locale, checked case by case against Excel. */
import type { Cx, FnDef } from '../../args.js'
import { E, type CErr, type V } from '../../values.js'

function value(cx: Cx, v: Exclude<V, CErr>): V {
  if (typeof v === 'number') return v
  if (v === null) return 0
  if (typeof v === 'boolean') return E.VALUE
  return cx.rules.valueText(v) ?? E.VALUE
}

/**
 * Group separators go first (after the decimal one they are #VALUE!), the decimal separator becomes the
 * locale's and the rest reads as in VALUE; whitespace anywhere is dropped and each trailing `%` divides by 100.
 * The separators default to the locale's.
 */
function numbervalue(cx: Cx, text: string, decimal?: string, group?: string): V {
  if (decimal === '' || group === '') return E.VALUE
  const { rules } = cx
  const d = decimal?.[0] ?? rules.decimal
  // the locale's group separator as the decimal one makes its decimal separator the default group separator
  const g = group?.[0] ?? (d === rules.group ? rules.decimal : rules.group)
  if (d === g) return E.VALUE
  let s = text.replace(/[ \t\n\r]/g, '')
  if (s === '') return 0
  let pct = 0
  for (; s.endsWith('%'); pct++) s = s.slice(0, -1)
  const at = s.indexOf(d)
  const frac = at < 0 ? '' : s.slice(at + 1)
  if (frac.includes(d) || frac.includes(g)) return E.VALUE
  const n = rules.valueText((at < 0 ? s : s.slice(0, at)).split(g).join('') + (at < 0 ? '' : rules.decimal + frac))
  return n === null ? E.VALUE : n / 100 ** pct
}

export const functions: Record<string, FnDef> = {
  NUMBERVALUE: { args: 's s? s?', cx: true, fn: numbervalue },
  VALUE: { args: 'v', cx: true, fn: value },
}
