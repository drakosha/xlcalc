/**
 * What of a slicer and a timeline Excel opens and keeps from the file `writeXlsx` writes, each case the file itself. A session's new
 * slicer (`editStructure` `addSlicer`, `writeXlsx` of `getSnapshot`) has a place, and its frame cost the file on Excel for Mac without
 * the defined name of its cache, which the stored sets of PivotTables did not see, as they read books Excel saved. The writer now adds
 * the name where the book holds none, and leaves out what has no place, as Excel loses it on saving. And the file without the name, for
 * each of a slicer and a timeline, so what the writer adds is what stands between Excel and the refusal.
 *
 * `refused` and `SAVED` were committed before Excel ran this set. The model: Excel asks the cache's name of the workbook's names, in any
 * case and with any value there, a sheet's own not counting, for a frame alone; so all of the writer's files open, each saved with
 * what it holds, and the file without the name is refused for a slicer and for a timeline alike.
 */
import { readFileSync } from 'node:fs'
import type { BookInput, SheetInput, StructureOp } from '../../../src/contract.js'
import { createSession, editStructure, registerPivot } from '../../../src/index.js'
import { pivotEngine } from '../../../src/pivot/index.js'
import { readXlsx, writeXlsx } from '../../../src/xlsx.js'
import type { LinksCase, LinksCases } from './cases.js'
import { REFUSED, subs } from './file-refusals.js'

const day = (y: number, m: number, d: number) => Date.UTC(y, m - 1, d) / 86_400_000 + 25_569
const SOURCE: [string, number, string, number][] = [
  ['North', day(2024, 1, 5), 'Apple', 10],
  ['South', day(2024, 6, 1), 'Banana', 20],
  ['East', day(2025, 2, 1), 'Apple', 5],
  ['North', day(2025, 3, 9), 'Cherry', 7],
  ['West', day(2025, 11, 30), 'Banana', 1],
]
const cells: SheetInput['cells'] = { A1: { value: 'Region' }, B1: { value: 'Date' }, C1: { value: 'Product' }, D1: { value: 'Amount' } }
SOURCE.forEach(([region, date, product, amount], i) => {
  cells[`A${i + 2}`] = { value: region }
  cells[`B${i + 2}`] = { value: date, numberFormat: 'yyyy-mm-dd' }
  cells[`C${i + 2}`] = { value: product }
  cells[`D${i + 2}`] = { value: amount }
})
const location = { range: 'A3:B4', firstHeaderRow: 1, firstDataRow: 1, firstDataColumn: 1 }
const SHEETS: SheetInput[] = [
  { name: 'D', cells },
  { name: 'P', cells: {}, pivotTables: [{ name: 'PT1', cache: 1, location, fields: [{}, {}, {}, {}], rows: [0], values: [{ field: 3 }] }] },
]
const CACHES: Omit<BookInput, 'sheets'> = { pivotCaches: [{ id: 1, source: { type: 'range', sheet: 'D', range: 'A1:D6' }, fields: ['Region', 'Date', 'Product', 'Amount'].map((name) => ({ name })) }] }
const PT = { sheet: 'P', name: 'PT1' }
const add = (field: string, name?: string): StructureOp => ({ op: 'addSlicer', sheet: 'P', slicer: name ? { name } : {}, cache: { field, pivots: [PT] } })
const TIMELINE: StructureOp = { op: 'addTimeline', sheet: 'P', timeline: {}, cache: { field: 'Date', pivots: [PT] } }
const WORKBOOK = 'xl/workbook.xml'
const noName = (name: string) => subs([WORKBOOK, new RegExp(`<definedName name="${name}">#N/A</definedName>`), ''])

/** A book through a session, as a host makes it: the operations after a refresh, its snapshot written. */
export async function sessionFile(book: BookInput, ops: StructureOp[]) {
  registerPivot(pivotEngine)
  const s = createSession(book)
  for (const op of [{ op: 'refreshPivots' } as StructureOp, ...ops]) {
    const r = s.editStructure(op)
    if (!r.ok) throw new Error(r.message)
  }
  const written = await writeXlsx(s.getSnapshot())
  s.close()
  return written
}

/** A file read and written anew, the mirror sheets of `book` beside its sheets. */
export async function againFile(data: Uint8Array, book: BookInput) {
  const read = (await readXlsx(data)).book
  return writeXlsx({ book: { ...read, sheets: [...read.sheets, ...book.sheets.filter((s) => /^W[TNX]$/.test(s.name) && !read.sheets.some((r) => r.name === s.name))] } })
}

/** The book of the cases written from a book: refreshed, with what `change` does, each cache its slicer or timeline put in by hand. */
function given(change: (b: BookInput) => void): BookInput {
  registerPivot(pivotEngine)
  const r = editStructure({ sheets: SHEETS, ...CACHES }, { op: 'refreshPivots' })
  if (!r.ok) throw new Error(r.message)
  const b = structuredClone(r.book) as BookInput
  change(b)
  return b
}
const slicerOf = (b: BookInput, field: string, name: string, cache: string, place?: { from: string; to: string }) => {
  b.slicerCaches = [...(b.slicerCaches ?? []), ...(b.slicerCaches?.some((c) => c.name === cache) ? [] : [{ name: cache, field, pivots: [PT], cache: 1 }])]
  b.sheets[1].slicers = [...(b.sheets[1].slicers ?? []), { name, cache, ...(place ? { place } : {}) }]
}
const timelineOf = (b: BookInput, place?: { from: string; to: string }) => {
  b.timelineCaches = [{ name: 'NativeTimeline_Date', field: 'Date', pivots: [PT], cache: 1 }]
  b.sheets[1].timelines = [{ name: 'Date', cache: 'NativeTimeline_Date', ...(place ? { place } : {}) }]
}
const HERE = { from: 'H2', to: 'K12' }
const THERE = { from: 'M2', to: 'P12' }

const base = (id: string, about: string, e: Partial<LinksCase>): LinksCase => ({ id, about, sheets: SHEETS, linked: {}, links: [], read: [{ sheet: 'D', range: 'D2' }], predicted: { 'D!D2': 10 }, written: true, book: CACHES, ...e })
const session = (id: string, about: string, ops: StructureOp[], e: Partial<LinksCase> = {}) => base(id, about, { file: async (b) => (await sessionFile(b, ops)).data, ...e })
const book = (id: string, about: string, change: (b: BookInput) => void, e: Partial<LinksCase> = {}) =>
  base(id, about, {
    file: async (b) => {
      const made = given(change)
      return (await writeXlsx({ book: { ...made, sheets: [...made.sheets, ...b.sheets.filter((s) => /^W[TNX]$/.test(s.name))] } })).data
    },
    ...e,
  })
const excelFile = (path: string) => new Uint8Array(readFileSync(new URL(`../../../test/excel/win/pivot/${path}`, import.meta.url)))
const again = (id: string, about: string, path: string) =>
  base(id, about, {
    file: async (b) => (await againFile(excelFile(path), b)).data,
    sheets: [{ name: 'D', cells: {} }],
    book: undefined,
    predicted: undefined,
  })

/** What the file Excel saved holds, predicted: the parts of each kind and its defined names but its filter databases. */
export const SAVED: Record<string, { slicers: number; timelines: number; names: string[] }> = {
  'session-slicer': { slicers: 1, timelines: 0, names: ['Slicer_Region'] },
  'session-timeline': { slicers: 0, timelines: 1, names: ['NativeTimeline_Date'] },
  'session-both': { slicers: 1, timelines: 1, names: ['NativeTimeline_Date', 'Slicer_Region'] },
  'session-two': { slicers: 2, timelines: 0, names: ['Slicer_Product', 'Slicer_Region'] },
  'session-taken': { slicers: 1, timelines: 0, names: ['Slicer_Region', 'Slicer_Region1'] },
  'session-again': { slicers: 1, timelines: 1, names: ['NativeTimeline_Date', 'Slicer_Region'] },
  'no-place-slicer': { slicers: 0, timelines: 0, names: [] },
  'no-place-timeline': { slicers: 0, timelines: 0, names: [] },
  'no-place-cache-beside': { slicers: 1, timelines: 0, names: ['Slicer_Region'] },
  'no-place-same-cache': { slicers: 1, timelines: 0, names: ['Slicer_Region'] },
  'placed-apart': { slicers: 2, timelines: 0, names: ['Slicer_Product', 'Slicer_Region'] },
  'book-name': { slicers: 1, timelines: 0, names: ['Slicer_Region'] },
  'book-name-case': { slicers: 1, timelines: 0, names: ['SLICER_REGION'] },
  'book-name-other': { slicers: 1, timelines: 0, names: ['Slicer_Region'] },
  'book-name-sheet': { slicers: 1, timelines: 0, names: ['Slicer_Region', 'Slicer_Region'] },
  'excel-slicer': { slicers: 1, timelines: 0, names: ['Slicer_Region'] },
  'excel-timeline': { slicers: 0, timelines: 1, names: ['Timeline_Date'] },
}

export function buildLinksCases(): LinksCases {
  const slicer = add('Region')
  const twice = async (b: BookInput) => {
    const first = await sessionFile(b, [slicer, TIMELINE])
    return (await againFile(first.data, b)).data
  }
  return {
    cases: [
      session('session-slicer', 'addSlicer through a session and its snapshot', [slicer]),
      session('session-timeline', 'addTimeline through a session and its snapshot', [TIMELINE]),
      session('session-both', 'a slicer and a timeline through a session', [slicer, TIMELINE]),
      session('session-two', 'two slicers of two fields through a session', [slicer, add('Product')]),
      session('session-taken', 'a slicer through a session in a book with the name it would take', [slicer], { names: [{ name: 'Slicer_Region', ref: 'D!$A$1' }] }),
      base('session-again', 'the session’s file read and written anew: one name of each', { file: twice }),
      book('no-place-slicer', 'a slicer without a place, in a book', (b) => slicerOf(b, 'Region', 'Region', 'Slicer_Region')),
      book('no-place-timeline', 'a timeline without a place, in a book', (b) => timelineOf(b)),
      book('no-place-cache-beside', 'a slicer with a place and another without, of two caches', (b) => {
        slicerOf(b, 'Region', 'Region', 'Slicer_Region', HERE)
        slicerOf(b, 'Product', 'Product', 'Slicer_Product')
      }),
      book('no-place-same-cache', 'a slicer with a place and another without, of one cache', (b) => {
        slicerOf(b, 'Region', 'Region', 'Slicer_Region', HERE)
        slicerOf(b, 'Region', 'Region 1', 'Slicer_Region')
      }),
      book('book-name', 'a slicer with a place, the book holding the name', (b) => {
        slicerOf(b, 'Region', 'Region', 'Slicer_Region', HERE)
        b.names = [{ name: 'Slicer_Region', ref: '#N/A' }]
      }),
      book('book-name-case', 'a slicer with a place, the book holding the name in another case', (b) => {
        slicerOf(b, 'Region', 'Region', 'Slicer_Region', HERE)
        b.names = [{ name: 'SLICER_REGION', ref: '#N/A' }]
      }),
      book('book-name-other', 'a slicer with a place, the book holding the name for a cell', (b) => {
        slicerOf(b, 'Region', 'Region', 'Slicer_Region', HERE)
        b.names = [{ name: 'Slicer_Region', ref: 'D!$A$1' }]
      }),
      book('book-name-sheet', 'a slicer with a place, a sheet holding the name as its own', (b) => {
        slicerOf(b, 'Region', 'Region', 'Slicer_Region', HERE)
        b.names = [{ name: 'Slicer_Region', ref: 'D!$A$1', scope: 'P' }]
      }),
      again('excel-slicer', 'a file of Excel’s with a slicer read and written anew', 'slicer/row-two.xlsx'),
      again('excel-timeline', 'a file of Excel’s with a timeline read and written anew', 'slicer/timeline-range.xlsx'),
      session('no-name-slicer', 'the session’s slicer, the name taken out of its file', [slicer], { edit: noName('Slicer_Region'), ...REFUSED }),
      session('no-name-timeline', 'the session’s timeline, the name taken out of its file', [TIMELINE], { edit: noName('NativeTimeline_Date'), ...REFUSED }),
      book('placed-apart', 'two slicers with places, of two caches, in a book', (b) => {
        slicerOf(b, 'Region', 'Region', 'Slicer_Region', HERE)
        slicerOf(b, 'Product', 'Product', 'Slicer_Product', THERE)
      }),
    ],
  }
}
