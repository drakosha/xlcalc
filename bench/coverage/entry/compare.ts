/**
 * `parseEntry`'s answer against what Excel made of the same text: refused or taken alike, and a taken text's value or
 * formula and format equal to the cell Excel saved. Formats compare as the xlsx reader reads them: a built-in format by
 * its number (a code equal to one that every locale shows alike is that number), General as none. Numbers compare to 15
 * significant digits, as Excel writes them.
 */
import type { CellInput, EntryResult, NumberFormat } from '../../../src/contract.js'
import { BUILTIN_ID } from '../../../src/xlsx/styles/numfmt.js'
import type { EntryAnswer } from './cases.js'

/** A format as the xlsx reader gives it: none for General, a built-in format's number where a file stores that number. */
export const readFormat = (f: NumberFormat | undefined): NumberFormat | undefined => (f === undefined || f === 0 || (typeof f === 'string' && f.toLowerCase() === 'general') ? undefined : typeof f === 'string' ? (BUILTIN_ID.get(f) ?? f) : f)

const same = (a: unknown, b: unknown) => (typeof a === 'number' && typeof b === 'number' ? a === b || Number(a.toPrecision(15)) === Number(b.toPrecision(15)) : a === b)

/** The parts of a cell that are compared: its value or formula, and its format as read. */
const held = (c: CellInput | null) => ({ value: c?.formula === undefined ? (c?.value ?? null) : null, formula: c?.formula ?? null, format: readFormat(c?.format) ?? null })

/** The differences of `parseEntry`'s answer from Excel's, as text; none when they agree. */
export function entryDifferences(got: EntryResult, want: EntryAnswer): string[] {
  if (want.refused !== undefined) return got.ok ? [`Excel refused it (${want.refused}), the engine took it as ${JSON.stringify(got.input)}`] : []
  if (!got.ok) return [`Excel took it as ${JSON.stringify(want.cell)}, the engine refused it: ${got.code} ${got.message} at ${got.at}`]
  const [g, w] = [held(got.input), held(want.cell)]
  return (['value', 'formula', 'format'] as const).filter((k) => !same(g[k], w[k])).map((k) => `${k}: Excel ${JSON.stringify(w[k])}, the engine ${JSON.stringify(g[k])}`)
}
