/** BIN2DEC … OCT2HEX: ten-digit two's complement numbers with Excel's checks. This engine's own. */
import { generalText } from '../../format.js'
import { E, isErr, type CErr, type V } from '../../values.js'
import type { FnDef } from '../../args.js'
import { atp, need } from '../date/atp.js'

const DIGITS = 10

/** The number's text: a number in the General format, a blank empty (0); a boolean is #VALUE!. */
function digitsOf(v: V): string | CErr {
  if (typeof v === 'boolean') return E.VALUE
  return v === null ? '' : typeof v === 'number' ? generalText(v) : (v as string)
}

/** Ten digits at most in radix r (letters of either case), negative from r^10 / 2 up; else #NUM!. */
function read(v: V, r: number): number | CErr {
  const s = digitsOf(v)
  if (isErr(s)) return s
  if (s.length > DIGITS) return E.NUM
  let n = 0
  for (const ch of s) {
    const d = parseInt(ch, 36)
    if (!(d < r)) return E.NUM
    n = n * r + d
  }
  const full = r ** DIGITS
  return s.length === DIGITS && n >= full / 2 ? n - full : n
}

/** n in radix r, upper case; a negative n takes all ten digits (`places`, 1 to 10, is checked as it is read). */
function write(n: number, r: number, places?: number): string | CErr {
  const full = r ** DIGITS
  if (n < -full / 2 || n >= full / 2) return E.NUM
  if (n < 0) return (full + n).toString(r).toUpperCase()
  const s = n.toString(r).toUpperCase()
  if (places === undefined) return s
  return s.length > places ? E.NUM : s.padStart(places, '0')
}

/** Excel reads `places` first: a boolean is #VALUE!, a count outside 1 to 10 #NUM!, before the number is looked at. */
const PLACES = { first: [1], check: { 1: need((p: number) => p >= 1 && p <= DIGITS) } }

const RADIX: Record<string, number> = { BIN: 2, OCT: 8, DEC: 10, HEX: 16 }

export const functions: Record<string, FnDef> = {}
for (const [from, rf] of Object.entries(RADIX))
  for (const [to, rt] of Object.entries(RADIX)) {
    if (from === to) continue
    const name = `${from}2${to}`
    if (from === 'DEC') functions[name] = atp('x j?', (n: number, places?: number) => write(Math.trunc(n), rt, places), PLACES)
    else if (to === 'DEC') functions[name] = { args: 'v', fn: (v: V) => read(v, rf) }
    else
      functions[name] = atp(
        'v j?',
        (v: V, places?: number) => {
          const n = read(v, rf)
          return isErr(n) ? n : write(n, rt, places)
        },
        PLACES,
      )
  }
