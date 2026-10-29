/**
 * `parseEntry`'s answer against what Excel made of the same text: refused or taken alike, a refused text with the
 * correction Excel's window offered, and a taken text's value or formula and format equal to the cell Excel saved. Formulas compare by their canonical text (Excel keeps the spaces
 * typed). Formats compare by what they show in the locale: a few values shown by each (Excel saves a built-in money format
 * by its code, `d mmmm yyyy` as `d\ mmmm\ yyyy`). Numbers compare exactly: the value is the one stored in the cell, which
 * Excel writes to the file with every digit (`1.255%` is 0.01255, not 1.255/100).
 */
import type { CellInput, EntryResult, NumberFormat } from '../../../src/contract.js'
import { createSession, displayValue, parseEntry, validateFormula } from '../../../src/index.js'
import { numberFormatOf } from '../stored.js'
import { entryBook, type EntryCase, type EntryData } from './cases.js'

/**
 * The engine's answer to a case in the locale on the day `now`: `parseEntry` into a cell of the case's format; for a set
 * with its own data, a session's `previewCellText` in A1 of the case's book, whose formulas read the formats of the data.
 */
export function engineEntry(c: EntryCase, tag: string, now: string, data?: EntryData): EntryResult {
  if (!data) return parseEntry(c.text, { locale: tag, now, ...(c.format !== undefined ? { cell: { numberFormat: c.format } } : {}) })
  const s = createSession({ ...entryBook([c], data), locale: tag }, { clock: { now: `${now}T12:00:00Z`, seed: 1 } })
  try {
    const r = s.previewCellText('S', 'A1', c.text)
    return r.ok ? { ok: true, input: r.input } : ({ at: 0, ...r } as EntryResult)
  } finally {
    s.close()
  }
}

const SAMPLES = [46096.770833, -1234.5678, 0.05, 0, 1.5]

/** What a format shows in the locale: the samples shown by it in a wide column; null for General. */
export function formatKey(f: NumberFormat | undefined, tag: string): string | null {
  if (f === undefined || f === 0 || (typeof f === 'string' && f.toLowerCase() === 'general')) return null
  return SAMPLES.map((v) => displayValue(v, { numberFormat: f, locale: tag, width: 60 }).text).join(' | ')
}

const canonical = (f: string) => {
  const v = validateFormula(f)
  return v.ok ? v.canonical : f
}

const same = (a: unknown, b: unknown) => a === b

/** The parts of a cell that are compared: its value or formula, and its format as the locale shows it. */
export const held = (c: CellInput | null, tag: string) => ({
  value: c?.formula === undefined ? (c?.value ?? null) : null,
  formula: c?.formula === undefined ? null : canonical(c.formula),
  format: formatKey(numberFormatOf(c), tag),
})

/** The differences of `parseEntry`'s answer to `text` in the locale `tag` from Excel's, as text; none when they agree. */
export function entryDifferences(got: EntryResult, want: { refused?: string; cell: CellInput | null }, tag: string, text: string): string[] {
  if (want.refused !== undefined) {
    if (got.ok) return [`Excel refused it (${want.refused.split('\n')[0]}), the engine took it as ${JSON.stringify(got.input)}`]
    // the formula Excel's window offers to correct the text to, between the marks of left-to-right text it puts around it;
    // only one that closes what the text left open (the spaces after = dropped) is the engine's to offer
    const offered = /\u202d(.*?)\u202c/s.exec(want.refused)?.[1]
    const typed = text[0] + text.slice(1).replace(/^\s+/, '')
    const fix = offered !== undefined && offered.startsWith(typed) && /^["')}]+$/.test(offered.slice(typed.length)) ? offered : null
    return fix === (got.correction ?? null) ? [] : [`correction: Excel ${JSON.stringify(offered ?? null)}, the engine ${JSON.stringify(got.correction ?? null)}`]
  }
  if (!got.ok) return [`Excel took it as ${JSON.stringify(want.cell)}, the engine refused it: ${got.code} ${got.message} at ${got.at}`]
  const [g, w] = [held(got.input, tag), held(want.cell, tag)]
  const raw = (c: CellInput | null, k: 'value' | 'formula' | 'format') => JSON.stringify((k === 'value' && c?.formula !== undefined ? null : k === 'format' ? numberFormatOf(c) : c?.[k]) ?? null)
  return (['value', 'formula', 'format'] as const).filter((k) => !same(g[k], w[k])).map((k) => `${k}: Excel ${raw(want.cell, k)}, the engine ${raw(got.input, k)}`)
}
