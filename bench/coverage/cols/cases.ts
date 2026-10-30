/**
 * Overlapping `<col>` elements, as a host's writer leaves them: each sheet's columns in order of `min`, a later `<col>`
 * wholly inside an earlier one, differing in width, style, both or nothing. Whether Excel opens such a file without
 * repairing it, which width and style a column inside takes, and what Excel saves. Each case is a book of its own, so
 * one Excel repairs does not hide the others.
 */
import type { BookInput } from '../../../src/contract.js'
import { indexToCol } from '../../../src/model.js'
import { writeXlsx } from '../../../src/xlsx.js'
import { openArchive } from '../../../src/xlsx/zip/read.js'
import { rawFile } from '../xlsx/sent.js'

/** The two styles a case's `<col>` may give: `P` a percent, `Q` two decimals (their xf ids are found in the file written). */
export type Style = 'P' | 'Q'

/** A `<col>`: its first and last column (1-based, as xlsx), its width and its style. */
export type Col = [min: number, max: number, width: number, style?: Style]

export interface ColsCase {
  id: string
  cols: Col[]
  /** The columns read (1-based): their width and the number format of their first cell. */
  read: number[]
}

const upTo = (n: number) => Array.from({ length: n }, (_, i) => i + 1)

/** The form of the files: an earlier `<col>` over A:J and a later one over C inside it. */
export const CASES: ColsCase[] = [
  { id: 'apart', cols: [[1, 2, 9], [3, 3, 20], [4, 10, 9]], read: upTo(11) },
  { id: 'width', cols: [[1, 10, 9], [3, 3, 20]], read: upTo(11) },
  { id: 'style-width', cols: [[1, 10, 9, 'P'], [3, 3, 20, 'Q']], read: upTo(11) },
  { id: 'style', cols: [[1, 10, 9, 'P'], [3, 3, 9, 'Q']], read: upTo(11) },
  { id: 'same', cols: [[1, 10, 9, 'P'], [3, 3, 9, 'P']], read: upTo(11) },
  { id: 'style-only-inner', cols: [[1, 10, 9], [3, 3, 9, 'Q']], read: upTo(11) },
  { id: 'edges', cols: [[2, 10, 9, 'P'], [2, 2, 14], [5, 6, 16, 'Q'], [10, 10, 18]], read: upTo(11) },
  { id: 'many', cols: [[1, 60, 9, 'P'], ...Array.from({ length: 24 }, (_, i): Col => [2 + 2 * i, 2 + 2 * i, 10 + i, i % 2 ? 'Q' : undefined])], read: upTo(61) },
  // not in the host's files: a later `<col>` that starts inside an earlier one and ends past it
  { id: 'partial', cols: [[1, 5, 9, 'P'], [4, 8, 20, 'Q']], read: upTo(9) },
]

/** The book under the `<cols>`: a value in K1, and columns M and N with the two styles, so the styles part holds both. */
const book = (): BookInput => ({
  sheets: [{ name: 'S', cells: { K1: { value: 'cols' } }, columns: { M: { numberFormat: '0%' }, N: { numberFormat: '0.00' } } }],
})

/** The xf ids the writer gave columns M and N. */
function styleIds(sheet: string): Record<Style, string> {
  const at = (n: number) => new RegExp(`<col min="${n}" max="${n}"[^>]*? style="(\\d+)"`).exec(sheet)?.[1]
  const [p, q] = [at(13), at(14)]
  if (!p || !q) throw new Error(`no styles for M and N in ${sheet.slice(0, 400)}`)
  return { P: p, Q: q }
}

/** A `<col>` as the host's writer writes it. */
const colXml = ([min, max, width, style]: Col, ids: Record<Style, string>) => `<col min="${min}" max="${max}" width="${width}"${style ? ` style="${ids[style]}"` : ''} customWidth="1"/>`

/** The case's file: the book written, its `<cols>` replaced by the case's. */
export async function caseFile(c: ColsCase): Promise<Uint8Array> {
  const { data } = await writeXlsx({ book: book() })
  const sheet = (await openArchive(data).text('xl/worksheets/sheet1.xml'))!
  const cols = /<cols>.*?<\/cols>/.exec(sheet)![0]
  const ids = styleIds(sheet)
  return rawFile(data, [{ part: 'xl/worksheets/sheet1.xml', from: cols, to: `<cols>${c.cols.map((x) => colXml(x, ids)).join('')}</cols>` }])
}

/** The letters of the columns a case reads. */
export const letters = (c: ColsCase) => c.read.map((n) => indexToCol(n - 1))

/** What Excel read of a column: its width and the number format of its first cell. */
export type ColRead = [width: number, format: string]

/** A case's answer: Excel repaired the book (the window's text), or what it read of each column and the `<col>`s it saved. */
export type ColsAnswer = { repaired: string } | { read: Record<string, ColRead>; saved: string[] }

/** The `<col>` elements of the first sheet of a saved file, as written. */
export async function savedCols(saved: Uint8Array): Promise<string[]> {
  const sheet = (await openArchive(saved).text('xl/worksheets/sheet1.xml'))!
  return [...sheet.matchAll(/<col\s[^>]*\/>/g)].map((m) => m[0])
}
