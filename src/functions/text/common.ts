/** Helpers shared by the text functions. */
import type { Segment } from '../../args.js'
import type { Rules } from '../../locale/rules.js'
import { MAX_TEXT } from '../../model.js'
import { E, isErr, type CErr, type V } from '../../values.js'

export { MAX_TEXT }

export const capped = (s: string): V => (s.length > MAX_TEXT ? E.VALUE : s)

/** CONCAT and TEXTJOIN: a longer text is #CALC!, as in Excel (other functions answer #VALUE!). */
export const joined = (s: string): V => (s.length > MAX_TEXT ? E.CALC : s)

/** Delimiters as texts (TEXTJOIN, TEXTBEFORE, TEXTSPLIT): a range or array gives each of its values, an error is the answer. */
export function delimiters(seg: Segment, rules: Rules): string[] | CErr {
  const out: string[] = []
  for (const v of Array.isArray(seg) ? (seg as readonly V[]) : [seg as V]) {
    const t = rules.toText(v)
    if (isErr(t)) return t
    out.push(t)
  }
  return out
}
