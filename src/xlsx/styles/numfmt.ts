/**
 * Built-in number formats: a file stores only their number. The model keeps formats 0 to 49 by number (`NumberFormat`)
 * and shows them by the book's locale; the reader gives them as numbers. Formats 50 to 58 (East Asian dates) the model
 * does not have: the reader gives the code Excel for Mac en-US shows for them (measured: `test/excel/xlsx/formats.json`).
 * The writer stores a code equal to a built-in format that every locale shows alike by that format's number, as Excel
 * does, and the codes of the formats a locale shows otherwise as codes of their own (ids from 164), so the file shows
 * what `format` says in every locale.
 */
import { EN_BUILTIN, isBuiltin } from '../../numfmt/builtin.js'

/** The codes of the built-in formats 1 to 58 in en-US: Excel for Windows's to 49 (`numfmt/builtin.ts`), Excel for Mac's from 50. */
export const BUILTIN: Readonly<Record<number, string>> = Object.fromEntries(
  Array.from({ length: 58 }, (_, i) => i + 1).map((id) => [id, isBuiltin(id) ? EN_BUILTIN[id] : 'm/d/yy']),
)

const span = (lo: number, hi: number) => Array.from({ length: hi - lo + 1 }, (_, i) => lo + i)

/** The built-in formats Excel shows by the regional settings: money, dates and times, negatives in parentheses. */
const REGIONAL: ReadonlySet<number> = new Set([...span(5, 8), ...span(14, 44), ...span(50, 58)])

/** Built-in ids by code, for the formats a file can store by number in every locale. */
export const BUILTIN_ID: ReadonlyMap<string, number> = new Map(
  Object.entries(BUILTIN)
    .filter(([id]) => !REGIONAL.has(Number(id)))
    .map(([id, code]) => [code, Number(id)]),
)
