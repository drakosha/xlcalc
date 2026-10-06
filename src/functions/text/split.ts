/** TEXTBEFORE and TEXTAFTER, this engine's own (Excel 2024). */
import type { FnDef, Segment } from '../../args.js'
import { E, isErr, type V } from '../../values.js'
import { texts } from './common.js'

type Span = [number, number]

/** The earliest delimiter at or after `p`, the longest on a tie. */
function next(hay: string, keys: string[], p: number): Span | null {
  let best: Span | null = null
  for (const k of keys) {
    const at = hay.indexOf(k, p)
    if (at >= 0 && (!best || at < best[0] || (at === best[0] && at + k.length > best[1]))) best = [at, at + k.length]
  }
  return best
}

/** The latest delimiter ending at or before `p`, the longest on a tie. */
function prev(hay: string, keys: string[], p: number): Span | null {
  let best: Span | null = null
  for (const k of keys) {
    const at = p - k.length < 0 ? -1 : hay.lastIndexOf(k, p - k.length)
    if (at >= 0 && (!best || at + k.length > best[1] || (at + k.length === best[1] && at < best[0]))) best = [at, at + k.length]
  }
  return best
}

/** The `n`th occurrence without overlaps, from the start (n > 0) or the end; `matchEnd` counts the text's end (start) as one more. */
function occurrence(hay: string, keys: string[], n: number, matchEnd: boolean): Span | null {
  let p = n > 0 ? 0 : hay.length
  for (let k = 1; k <= Math.abs(n); k++) {
    const s = n > 0 ? next(hay, keys, p) : prev(hay, keys, p)
    if (!s) return matchEnd && k === Math.abs(n) ? (n > 0 ? [hay.length, hay.length] : [0, 0]) : null
    if (k === Math.abs(n)) return s
    // an empty delimiter occurs at every position
    p = n > 0 ? (s[1] > s[0] ? s[1] : s[1] + 1) : s[1] > s[0] ? s[0] : s[0] - 1
  }
  return null
}

const around = (after: boolean): FnDef => ({
  args: 's V i? i? i? a?',
  fn: (text: string, delims: Segment, n = 1, mode = 0, matchEnd = 0, notFound?: V): V => {
    const ds = texts(delims)
    if (isErr(ds)) return ds
    if (n === 0 || Math.abs(n) > text.length || (mode !== 0 && mode !== 1) || (matchEnd !== 0 && matchEnd !== 1)) return E.VALUE
    const s = mode ? occurrence(text.toLowerCase(), ds.map((d) => d.toLowerCase()), n, matchEnd === 1) : occurrence(text, ds, n, matchEnd === 1)
    if (!s) return notFound === undefined ? E.NA : notFound
    return after ? text.slice(s[1]) : text.slice(0, s[0])
  },
})

export const functions: Record<string, FnDef> = {
  TEXTAFTER: around(true),
  TEXTBEFORE: around(false),
}
