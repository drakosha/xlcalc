/**
 * Which section shows a value, as measured in Excel (`test/excel/cases/text-format.json`): without conditions
 * the first for positives (and zero with fewer than three sections), the second for negatives, the third for
 * zero, the fourth (or a section with `@`, which ending a code of two or three sections takes no number) for text;
 * with conditions the first section whose condition holds or that has none, else General when the code has one
 * section and nothing otherwise (TEXT answers #VALUE!, a cell shows `#`). A section writes its own minus when every
 * value it can receive is negative.
 */
import type { Cond, Section, Tok } from './parse.js'

/** The section of a number, or null for General; `signed` when the formatted number takes a minus. Null when no section takes it. */
interface Choice {
  sec: Section | null
  signed: boolean
}

const holds = (c: Cond, v: number): boolean =>
  c.op === '<' ? v < c.limit : c.op === '<=' ? v <= c.limit : c.op === '>' ? v > c.limit : c.op === '>=' ? v >= c.limit : c.op === '=' ? v === c.limit : v !== c.limit

/** Whether the conditional section `i` receives only negative numbers: no value from 0 up reaches it. */
function negativeOnly(secs: Section[], i: number): boolean {
  const limits = secs.slice(0, i + 1).flatMap((s) => (s.cond && s.cond.limit >= 0 ? [s.cond.limit] : []))
  const probes = [0, Number.MIN_VALUE, Number.MAX_VALUE, ...limits.flatMap((l) => [l, l * (1 + 2 ** -50) + Number.MIN_VALUE, l * (1 - 2 ** -50)])]
  return !probes.some((p) => p >= 0 && secs.slice(0, i).every((s) => !s.cond || !holds(s.cond, p)) && (!secs[i].cond || holds(secs[i].cond!, p)))
}

export function choose(secs: Section[], v: number): Choice | null {
  // a text section ending a code of two or three sections takes no number (`h:mm;@` shows a negative as `h:mm` does)
  const n = secs.length > 1 && secs.length < 4 && secs[secs.length - 1].toks.some((t) => t.t === 'text') ? secs.length - 1 : secs.length
  const ordered = secs.slice(0, Math.min(3, n))
  if (!ordered.some((s) => s.cond)) {
    if (v < 0 && n >= 2) return { sec: secs[1], signed: false }
    return { sec: v === 0 && n >= 3 ? secs[2] : secs[0], signed: true }
  }
  const i = ordered.findIndex((s) => !s.cond || holds(s.cond, v))
  if (i < 0) return secs.length === 1 ? { sec: null, signed: true } : null
  return { sec: ordered[i], signed: !negativeOnly(ordered, i) }
}

/** The section that shows text: the fourth, else one with `@`. */
export const textSection = (secs: Section[]): Section | undefined => (secs.length >= 4 ? secs[3] : secs.find((s) => s.toks.some((t) => t.t === 'text')))

/** What a section is for a number: General, a date, text (the number shows as General), digits, or literals alone. */
export function kind(toks: Tok[]): 'general' | 'date' | 'text' | 'number' | 'literal' {
  let k: 'general' | 'date' | 'text' | 'number' | 'literal' = 'literal'
  for (const t of toks) {
    if (t.t === 'general') return 'general'
    if (t.t === 'date' || t.t === 'elapsed' || t.t === 'ampm') k = 'date'
    else if (t.t === 'text' && k !== 'date') k = 'text'
    else if ((t.t === 'ph' || t.t === 'exp') && k === 'literal') k = 'number'
  }
  return k
}
