/**
 * `parseEntry`'s answer against what Excel made of the same text: refused or taken alike, and a taken text's value or
 * formula and format equal to the cell Excel saved. A format compares as the locale's Excel takes it: General as none,
 * a code that is a built-in format's code in the locale as that format's number (`h:mm` is 20), other codes by what they
 * show (Excel saves `d mmmm yyyy` as `d\ mmmm\ yyyy`). Numbers compare to 15 significant digits, as Excel writes them.
 */
import type { CellInput, EntryResult, NumberFormat } from '../../../src/contract.js'
import { builtinOf } from '../../../src/entry/forms.js'
import { EN_CODES } from '../../../src/numfmt.js'
import { parseFormat } from '../../../src/numfmt/parse.js'
import type { EntryAnswer } from './cases.js'

/** A format as the locale's Excel takes it: none for General, a built-in format's number, else the sections the code reads as. */
export function formatKey(f: NumberFormat | undefined, tag: string): string | number | null {
  if (f === undefined) return null
  const id = builtinOf(f, tag)
  if (id === 0) return null
  return id ?? JSON.stringify(parseFormat(f as string, EN_CODES) ?? f)
}

const same = (a: unknown, b: unknown) => (typeof a === 'number' && typeof b === 'number' ? a === b || Number(a.toPrecision(15)) === Number(b.toPrecision(15)) : a === b)

/** The parts of a cell that are compared: its value or formula, and its format as the locale takes it. */
const held = (c: CellInput | null, tag: string) => ({ value: c?.formula === undefined ? (c?.value ?? null) : null, formula: c?.formula ?? null, format: formatKey(c?.format, tag) })

/** The differences of `parseEntry`'s answer in the locale `tag` from Excel's, as text; none when they agree. */
export function entryDifferences(got: EntryResult, want: EntryAnswer, tag: string): string[] {
  if (want.refused !== undefined) return got.ok ? [`Excel refused it (${want.refused}), the engine took it as ${JSON.stringify(got.input)}`] : []
  if (!got.ok) return [`Excel took it as ${JSON.stringify(want.cell)}, the engine refused it: ${got.code} ${got.message} at ${got.at}`]
  const [g, w] = [held(got.input, tag), held(want.cell, tag)]
  return (['value', 'formula', 'format'] as const).filter((k) => !same(g[k], w[k])).map((k) => `${k}: Excel ${JSON.stringify(w[k])}, the engine ${JSON.stringify(g[k])}`)
}
