/** Helpers shared by the text functions. */
import type { Segment } from '../../args.js'
import { E, isErr, toText, type CErr, type V } from '../../values.js'

/** Excel's limit on the length of a text value. */
export const MAX_TEXT = 32_767

export const capped = (s: string): V => (s.length > MAX_TEXT ? E.VALUE : s)

/** CONCAT and TEXTJOIN: a longer text is #CALC!, as in Excel (other functions answer #VALUE!). */
export const joined = (s: string): V => (s.length > MAX_TEXT ? E.CALC : s)

/** Delimiters as texts (TEXTJOIN, TEXTBEFORE, TEXTSPLIT): a range or array gives each of its values, an error is the answer. */
export function delimiters(seg: Segment): string[] | CErr {
  const out: string[] = []
  for (const v of Array.isArray(seg) ? (seg as readonly V[]) : [seg as V]) {
    const t = toText(v)
    if (isErr(t)) return t
    out.push(t)
  }
  return out
}
