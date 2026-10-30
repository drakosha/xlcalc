/**
 * How Excel keeps a formula that begins with `#REF!`: a host's books hold `#REF!-R` with the value `#VALUE!`, which
 * Excel computes as `#REF!`. The book's sheet S has formulas over a text and a number in column A and on the sheet T,
 * which become `#REF!-…` when column A and T are deleted, and two formulas already `#REF!-B1` with values stored
 * otherwise (`#VALUE!`, 7). Excel reads them as the book opens, after a recalculation, after the deletion and after a
 * full calculation, and saves the book after the deletion. `npx tsx bench/coverage/errstore/mac.ts`.
 */
import type { BookInput } from '../../../src/contract.js'
import { writeXlsx } from '../../../src/xlsx.js'
import { openArchive } from '../../../src/xlsx/zip/read.js'
import { rawFile } from '../xlsx/sent.js'

export const FORMULAS = ['A1-B1', 'A2-B1', 'T!A1-B1', 'T!A2-B1', '-A1', 'A1&"x"', '#REF!-B1', '#REF!-B1', 'A1+B1', 'B1-A1']

/** The cells read before the deletion (C1:C10) and after it, one column to the left. */
export const BEFORE = FORMULAS.map((_, i) => `C${i + 1}`)
export const AFTER = FORMULAS.map((_, i) => `B${i + 1}`)

export const book = (): BookInput => ({
  sheets: [
    { name: 'S', cells: { A1: { value: 'abc' }, A2: { value: 3 }, B1: { value: 5 }, ...Object.fromEntries(FORMULAS.map((f, i) => [`C${i + 1}`, { formula: f }])) } },
    { name: 'T', cells: { A1: { value: 'abc' }, A2: { value: 3 } } },
  ],
})

/** The values stored for FORMULAS as a writer that computed them leaves them, but C7's `#VALUE!` and C8's 7. */
const STORED = ['e:#VALUE!', '-2', 'e:#VALUE!', '-2', 'e:#VALUE!', 's:abcx', 'e:#VALUE!', '7', 'e:#VALUE!', 'e:#VALUE!']

const cellXml = (f: string, v: string, i: number) => {
  const t = v.startsWith('e:') ? ' t="e"' : v.startsWith('s:') ? ' t="str"' : ''
  return `<c r="C${i + 1}"${t}><f>${f.replace(/&/g, '&amp;')}</f><v>${v.replace(/^[es]:/, '')}</v></c>`
}

/** The book with its formulas of the old kind and stored values, written by hand over what `writeXlsx` writes, without `fullCalcOnLoad`. */
export async function bookFile(): Promise<Uint8Array> {
  const { data } = await writeXlsx({ book: book() })
  const sheet = (await openArchive(data).text('xl/worksheets/sheet1.xml'))!
  const from = /<sheetData>.*<\/sheetData>/.exec(sheet)![0]
  const head: Record<number, string> = { 1: '<c r="A1" t="s"><v>0</v></c><c r="B1"><v>5</v></c>', 2: '<c r="A2"><v>3</v></c>' }
  const rows = FORMULAS.map((f, i) => `<row r="${i + 1}">${head[i + 1] ?? ''}${cellXml(f, STORED[i], i)}</row>`).join('')
  return rawFile(data, [
    { part: 'xl/worksheets/sheet1.xml', from, to: `<sheetData>${rows}</sheetData>` },
    // as Excel 365 saves a book, which it does not compute again as it opens
    { part: 'xl/workbook.xml', from: '<calcPr fullCalcOnLoad="1"/>', to: '<calcPr calcId="191029"/>' },
  ])
}

/** Each read: [formula, shown text] per cell. */
export type Read = [string, string][]
export interface ErrStoreAnswer {
  opened: Read
  recalculated: Read
  deleted: Read
  full: Read
  /** The saved file's cells of S, as written: `<c …>…</c>`. */
  saved: string[]
}
