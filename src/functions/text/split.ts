/** TEXTBEFORE and TEXTAFTER, this engine's own (Excel 2024). The delimiter may be several texts (an array or a range). */
import type { Cx, FnDef, Segment } from '../../args.js'
import { E, isErr, WHOLE, type V } from '../../values.js'
import { delimiters } from './common.js'

type Span = [number, number]

/** The first delimiter found from `p` on (to the left of `p` when `back`): its span, or null. */
function nearest(hay: string, keys: string[], p: number, back: boolean): Span | null {
  let best: Span | null = null
  for (const key of keys) {
    const at = back ? (p < key.length ? -1 : hay.lastIndexOf(key, p - key.length)) : hay.indexOf(key, p)
    if (at < 0) continue
    if (!best || (back ? at + key.length > best[1] : at < best[0])) best = [at, at + key.length]
  }
  return best
}

/** The `n`th occurrence without overlaps, from the start (n > 0) or the end; `matchEnd` counts the end (start) of the text as one more. */
function occurrence(hay: string, keys: string[], n: number, matchEnd: boolean): Span | null {
  // Excel finds an empty delimiter at the start (from the end: at the end) for any instance
  if (keys.includes('')) return n > 0 ? [0, 0] : [hay.length, hay.length]
  let p = n > 0 ? 0 : hay.length
  for (let k = 1; k <= Math.abs(n); k++) {
    const s = nearest(hay, keys, p, n < 0)
    if (!s) return matchEnd && k === Math.abs(n) ? (n > 0 ? [hay.length, hay.length] : [0, 0]) : null
    if (k === Math.abs(n)) return s
    p = n > 0 ? s[1] : s[0]
  }
  return null
}

const around = (after: boolean): FnDef => ({
  args: 's V n? i? i? a?',
  cx: true,
  fn: (cx: Cx, text: string, delim: Segment, instance?: number, mode = 0, matchEnd = 0, notFound?: V): V => {
    // the instance is rounded down, as Excel does (-1.2 is -2), a hair under a whole number taken for it
    const n = instance === undefined ? 1 : Math.floor(instance + WHOLE)
    const keys = delimiters(delim, cx.rules)
    if (isErr(keys)) return keys
    if (n === 0 || (mode !== 0 && mode !== 1) || (matchEnd !== 0 && matchEnd !== 1)) return E.VALUE
    const empty = keys.includes('')
    // empty text: an empty delimiter is found, else Excel answers #N/A only to the plain call
    if (text === '') return empty ? '' : n === 1 && !matchEnd && notFound === undefined ? E.NA : E.VALUE
    // an instance past the text's length is #VALUE!, and so are delimiters all longer than the text once an instance is given
    if (!empty && (Math.abs(n) > text.length || (instance !== undefined && keys.every((k) => k.length > text.length)))) return E.VALUE
    const s = mode ? occurrence(text.toLowerCase(), keys.map((k) => k.toLowerCase()), n, matchEnd === 1) : occurrence(text, keys, n, matchEnd === 1)
    // Excel returns the not-found value as text
    if (!s) return notFound === undefined ? E.NA : cx.rules.toText(notFound)
    return after ? text.slice(s[1]) : text.slice(0, s[0])
  },
})

export const functions: Record<string, FnDef> = {
  TEXTAFTER: around(true),
  TEXTBEFORE: around(false),
}
