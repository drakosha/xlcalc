/** BASE, DECIMAL, ROMAN (five forms) and ARABIC with Excel's limits. This engine's own. */
import { E, isErr, type CErr, type V } from '../../values.js'
import type { Cx, FnDef } from '../../args.js'
import type { Rules } from '../../locale/rules.js'
import { digitsValue } from '../common.js'

/** Excel's limit for the length of text arguments here. */
const MAX_LEN = 255

function base(n: number, radix: number, minLength = 0): string | CErr {
  const x = Math.trunc(n)
  const r = Math.trunc(radix)
  const len = Math.trunc(minLength)
  if (x < 0 || x >= 2 ** 53 || r < 2 || r > 36 || len < 0 || len > MAX_LEN) return E.NUM
  return x.toString(r).toUpperCase().padStart(len, '0')
}

/** DECIMAL: leading spaces are skipped, trailing ones are not (Excel: DECIMAL(" FF",16) is 255, "FF " #NUM!). */
function decimal(text: string, radix: number): number | CErr {
  const r = Math.trunc(radix)
  if (text.length > MAX_LEN) return E.VALUE
  if (r < 2 || r > 36) return E.NUM
  return digitsValue(text.replace(/^ +/, ''), r) ?? E.NUM
}

const NUMERALS: [string, number][] = [['M', 1000], ['D', 500], ['C', 100], ['L', 50], ['X', 10], ['V', 5], ['I', 1]]

/** Subtractive pairs by form: 0 is classic (IV, IX, XL, …), each further form adds pairs with a smaller first numeral. */
const PAIRS: [string, number][][] = [
  [['CM', 900], ['CD', 400], ['XC', 90], ['XL', 40], ['IX', 9], ['IV', 4]],
  [['LM', 950], ['LD', 450], ['VC', 95], ['VL', 45]],
  [['XM', 990], ['XD', 490], ['IC', 99], ['IL', 49]],
  [['VM', 995], ['VD', 495]],
  [['IM', 999], ['ID', 499]],
]

/** Roman numerals of a form, greedily from the largest symbol or pair (Excel's forms, checked for 32 numbers × 5 forms). */
function roman(n: number, form: number): string {
  const symbols = [...NUMERALS, ...PAIRS.slice(0, form + 1).flat()].sort((a, b) => b[1] - a[1])
  let out = ''
  for (let rest = n; rest > 0; ) {
    const [s, v] = symbols.find(([, v]) => v <= rest)!
    out += s
    rest -= v
  }
  return out
}

/** ROMAN's form: a number from 0 to 4; TRUE is 0 and FALSE 4. */
function formOf(v: V | undefined, rules: Rules): number | CErr {
  if (v === undefined || v === true) return 0
  if (v === false) return 4
  const f = rules.toNum(v)
  return isErr(f) ? f : Math.trunc(f)
}

function toRoman(cx: Cx, n: number, form?: V): string | CErr {
  const f = formOf(form, cx.rules)
  if (isErr(f)) return f
  const x = Math.trunc(n)
  return x < 0 || x > 3999 || f < 0 || f > 4 ? E.VALUE : roman(x, f)
}

const VALUE: Record<string, number> = Object.fromEntries(NUMERALS)

/** ARABIC: any sequence of numerals, a smaller one before a larger one subtracted; spaces around and a minus sign allowed. */
function arabic(text: string): number | CErr {
  if (text.length > MAX_LEN) return E.VALUE
  let s = text.trim().toUpperCase()
  const sign = s.startsWith('-') ? -1 : 1
  if (sign < 0) {
    s = s.slice(1)
    if (!s) return E.NUM
  }
  let n = 0
  for (let i = 0; i < s.length; i++) {
    const v = VALUE[s[i]]
    if (v === undefined) return E.VALUE
    const next = i + 1 < s.length ? VALUE[s[i + 1]] : undefined
    n += next !== undefined && v < next ? -v : v
  }
  return sign * n
}

export const functions: Record<string, FnDef> = {
  ARABIC: { args: 's', fn: arabic },
  BASE: { args: 'n n n?', fn: base },
  DECIMAL: { args: 's n', fn: decimal },
  ROMAN: { args: 'n v?', cx: true, fn: toRoman },
}
