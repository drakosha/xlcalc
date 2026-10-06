/** Helpers shared by the text functions. */
import type { Segment } from '../../args.js'
import { E, isErr, toText, type CErr, type V } from '../../values.js'

/** Excel's limit on the length of a text value. */
export const MAX_TEXT = 32_767

export const capped = (s: string): V => (s.length > MAX_TEXT ? E.VALUE : s)

/** A `V` argument as texts: a range gives each of its cells (TEXTJOIN, TEXTBEFORE take several delimiters). */
export function texts(seg: Segment): string[] | CErr {
  const out: string[] = []
  for (const v of Array.isArray(seg) ? (seg as readonly V[]) : [seg as V]) {
    const t = toText(v)
    if (isErr(t)) return t
    out.push(t)
  }
  return out
}
