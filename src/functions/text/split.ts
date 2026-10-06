/** TEXTBEFORE and TEXTAFTER, this engine's own (Excel 2024). The delimiter is one text: no array constants. */
import type { FnDef } from '../../args.js'
import { E, toText, type V } from '../../values.js'

type Span = [number, number]

/** The `n`th occurrence without overlaps, from the start (n > 0) or the end; `matchEnd` counts the end (start) of the text as one more. */
function occurrence(hay: string, key: string, n: number, matchEnd: boolean): Span | null {
  // Excel finds an empty delimiter at the start (from the end: at the end) for any instance
  if (key === '') return n > 0 ? [0, 0] : [hay.length, hay.length]
  let p = n > 0 ? 0 : hay.length
  for (let k = 1; k <= Math.abs(n); k++) {
    const at = n > 0 ? hay.indexOf(key, p) : p < key.length ? -1 : hay.lastIndexOf(key, p - key.length)
    if (at < 0) return matchEnd && k === Math.abs(n) ? (n > 0 ? [hay.length, hay.length] : [0, 0]) : null
    if (k === Math.abs(n)) return [at, at + key.length]
    p = n > 0 ? at + key.length : at
  }
  return null
}

const around = (after: boolean): FnDef => ({
  args: 's s i? i? i? a?',
  fn: (text: string, delim: string, n = 1, mode = 0, matchEnd = 0, notFound?: V): V => {
    if (n === 0 || (mode !== 0 && mode !== 1) || (matchEnd !== 0 && matchEnd !== 1)) return E.VALUE
    // empty text: an empty delimiter is found, else Excel answers #N/A only to the plain call
    if (text === '') return delim === '' ? '' : n === 1 && !matchEnd && notFound === undefined ? E.NA : E.VALUE
    if (Math.abs(n) > text.length && delim !== '') return E.VALUE
    const s = mode ? occurrence(text.toLowerCase(), delim.toLowerCase(), n, matchEnd === 1) : occurrence(text, delim, n, matchEnd === 1)
    // Excel returns the not-found value as text
    if (!s) return notFound === undefined ? E.NA : toText(notFound)
    return after ? text.slice(s[1]) : text.slice(0, s[0])
  },
})

export const functions: Record<string, FnDef> = {
  TEXTAFTER: around(true),
  TEXTBEFORE: around(false),
}
