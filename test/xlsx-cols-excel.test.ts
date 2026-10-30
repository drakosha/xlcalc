/**
 * Overlapping `<col>` elements against Excel for Mac and Windows (`bench/coverage/cols`): both open such files without
 * repairing them, a column takes the whole of the last `<col>` over it, and the file they save splits the earlier ones
 * around the later. The reader reads each case's columns as Excel saved them; `writeKept` leaves the `<cols>` of a file
 * whose columns did not change as they were and writes them without overlaps once they change.
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { CASES, caseFile, type ColsAnswer, letters } from '../bench/coverage/cols/cases.js'
import { partsHash } from '../bench/coverage/xlsx/sent.js'
import type { ColumnInput } from '../src/contract.js'
import { lineAt, overlaid } from '../src/layout/spans.js'
import { readXlsx } from '../src/xlsx.js'
import { openKept, writeKept } from '../src/xlsx-keep.js'
import { formatAreas } from '../src/xlsx/keep/formats.js'
import { BUILTIN } from '../src/xlsx/styles/numfmt.js'
import { openArchive } from '../src/xlsx/zip/read.js'

type Run = { sent: Record<string, string>; answers: Record<string, ColsAnswer> }
const run = (path: string) => JSON.parse(readFileSync(new URL(path, import.meta.url), 'utf8')) as Run
const RUNS = { Mac: run('./excel/cols/overlap.json'), Windows: run('./excel/win/cols/overlap.json') }

/** Each column's width as saved `<col>` elements give it, by column number. */
function savedWidths(cols: string[]): Map<number, number> {
  const out = new Map<number, number>()
  for (const c of cols) {
    const [min, max, width] = ['min', 'max', 'width'].map((k) => Number(new RegExp(` ${k}="([^"]*)"`).exec(c)![1]))
    for (let n = min; n <= max; n++) out.set(n, width)
  }
  return out
}

const sheetText = async (data: Uint8Array) => (await openArchive(data).text('xl/worksheets/sheet1.xml'))!
const colsOf = async (data: Uint8Array) => /<cols>.*?<\/cols>/.exec(await sheetText(data))?.[0]

describe.each(Object.entries(RUNS))('overlapping <col> elements, Excel for %s', (_, stored) => {
  it.each(CASES)('$id', async (c) => {
    const data = await caseFile(c)
    expect(await partsHash(data)).toBe(stored.sent[c.id])
    const answer = stored.answers[c.id]
    if (!('read' in answer)) throw new Error(`Excel repaired ${c.id}: ${answer.repaired}`)
    const read = await readXlsx(data)
    expect(read.issues).toEqual([])
    const widths = savedWidths(answer.saved)
    letters(c).forEach((l, i) => {
      const col: ColumnInput | undefined = lineAt(read.book.sheets[0].columns, 'c', c.read[i] - 1)
      expect([l, col?.width]).toEqual([l, widths.get(c.read[i])])
      // K1 holds a cell of its own, which shows its own format
      if (l !== 'K') expect([l, typeof col?.numberFormat === 'number' ? BUILTIN[col.numberFormat] : (col?.numberFormat ?? 'General')]).toEqual([l, answer.read[l][1]])
    })
  })
})

describe('writeKept over overlapping <col> elements', () => {
  it('keeps them as they were while the columns do not change, and writes them without overlaps once they do', async () => {
    const data = await caseFile(CASES.find((c) => c.id === 'style-width')!)
    const kept = await openKept(data)
    const book = kept.book
    const same = await writeKept(data, { book: { ...book, sheets: book.sheets.map((s) => ({ ...s, cells: { ...s.cells, K2: { value: 1 } } })) } })
    expect(await colsOf(same.data)).toBe(await colsOf(data))
    const wider = { ...book, sheets: book.sheets.map((s) => ({ ...s, columns: { ...s.columns, C: { ...s.columns!.C, width: 25 } } })) }
    const written = await colsOf((await writeKept(data, { book: wider })).data)
    expect(written).toBe('<cols><col min="1" max="2" width="9" customWidth="1" style="1"/><col min="3" max="3" width="25" customWidth="1" style="2"/><col min="4" max="10" width="9" customWidth="1" style="1"/></cols>')
  })
})

describe('overlaid', () => {
  it('gives each later span the lines it covers, cutting the earlier ones around it', () => {
    expect(overlaid([[0, 9, 'a'], [2, 2, 'b'], [4, 5, 'c'], [9, 9, 'd']])).toEqual([[0, 1, 'a'], [2, 2, 'b'], [3, 3, 'a'], [4, 5, 'c'], [6, 8, 'a'], [9, 9, 'd']])
    expect(overlaid([[0, 4, 'a'], [3, 7, 'b'], [1, 8, 'c']])).toEqual([[0, 0, 'a'], [1, 8, 'c']])
    const apart: [number, number, string][] = [[0, 1, 'a'], [3, 3, 'b']]
    expect(overlaid(apart)).toBe(apart)
  })
  it("gives writeKept's column formats the later <col>'s", () => {
    const xml = '<cols><col min="1" max="10" style="5"/><col min="3" max="3" style="6"/></cols>'
    const { areas } = formatAreas(xml, undefined, { qname: 'cols', local: 'cols', start: 0, open: 6, end: xml.length - 7 }, (s) => s === 5)
    expect(areas.map((a) => [a.rect.c1, a.rect.c2, a.s])).toEqual([[0, 1, 5], [3, 9, 5]])
  })
})
